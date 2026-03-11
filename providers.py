"""LLM provider interface and implementations for Clipboard Coach."""
import json
import os
import re
import sys
import time
import logging
from abc import ABC, abstractmethod

log = logging.getLogger("coach")

SYSTEM_PROMPT = """Communication coach. Analyze the message and respond in JSON only.
Fix grammar/spelling. Improve communication impact. Be direct. If already good, say so.

CRITICAL FORMATTING RULE: The rewrite MUST use the exact same structure as the original:
- If the original has numbered lists (1. 2. 3.), the rewrite MUST have numbered lists
- If the original has bullet points (- or *), the rewrite MUST have bullet points
- If the original has line breaks, the rewrite MUST have line breaks in the same places
- If the original has paragraphs, headings, or indentation, preserve them exactly
- Only change the WORDS, never the FORMAT

JSON format (use \\n for line breaks in rewrite):
{"verdict":"improve"|"good","issue":"2-4 word label"|null,"nudge":"one sentence","rewrite":"improved version with identical formatting"|null}"""


def _parse_response(content):
    """Extract JSON from model response text."""
    json_match = re.search(r"\{[\s\S]*\}", content)
    if not json_match:
        raise ValueError(f"Invalid response: {content[:200]}")
    return json.loads(json_match.group())


# ── Interface ──────────────────────────────────────────────────────────
class LLMProvider(ABC):
    """Base class for all LLM providers."""

    @abstractmethod
    def complete(self, system: str, user: str) -> tuple[str, float]:
        """Send a chat completion request.

        Returns:
            (response_text, api_duration_seconds)
        """

    def analyze(self, text: str, pattern_hint: str = "") -> tuple[dict, float]:
        """Analyze a message using the provider. Returns (result_dict, api_time)."""
        content, api_time = self.complete(SYSTEM_PROMPT + pattern_hint, text)
        return _parse_response(content), api_time

    @property
    @abstractmethod
    def display_name(self) -> str:
        """Human-readable name for logging."""


# ── Azure OpenAI ───────────────────────────────────────────────────────
class AzureOpenAIProvider(LLMProvider):
    """Azure OpenAI / Azure AI Foundry."""

    def __init__(self, endpoint: str, deployment: str, api_key: str,
                 api_version: str = "2025-01-01-preview"):
        from openai import AzureOpenAI
        self._deployment = deployment
        self._client = AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version=api_version,
        )

    def complete(self, system, user):
        chunks = []
        t0 = time.perf_counter()
        ttft = None
        stream = self._client.chat.completions.create(
            model=self._deployment,
            max_tokens=300,
            temperature=0,
            stream=True,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        for chunk in stream:
            if chunk.choices and chunk.choices[0].delta.content:
                if ttft is None:
                    ttft = time.perf_counter() - t0
                chunks.append(chunk.choices[0].delta.content)
        duration = time.perf_counter() - t0
        log.info("  [timing] API: %.2fs | first token: %.2fs",
                 duration, ttft or duration)
        return "".join(chunks).strip(), duration

    @property
    def display_name(self):
        return f"Azure OpenAI ({self._deployment})"


# ── OpenAI Direct ──────────────────────────────────────────────────────
class OpenAIProvider(LLMProvider):
    """OpenAI API directly (api.openai.com)."""

    def __init__(self, model: str = "gpt-4.1", api_key: str = None):
        from openai import OpenAI
        self._model = model
        self._client = OpenAI(api_key=api_key or os.environ.get("OPENAI_API_KEY", ""))

    def complete(self, system, user):
        chunks = []
        t0 = time.perf_counter()
        ttft = None
        stream = self._client.chat.completions.create(
            model=self._model,
            max_tokens=300,
            temperature=0,
            stream=True,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        for chunk in stream:
            if chunk.choices and chunk.choices[0].delta.content:
                if ttft is None:
                    ttft = time.perf_counter() - t0
                chunks.append(chunk.choices[0].delta.content)
        duration = time.perf_counter() - t0
        log.info("  [timing] API: %.2fs | first token: %.2fs",
                 duration, ttft or duration)
        return "".join(chunks).strip(), duration

    @property
    def display_name(self):
        return f"OpenAI ({self._model})"


# ── Anthropic ──────────────────────────────────────────────────────────
class AnthropicProvider(LLMProvider):
    """Anthropic Claude API."""

    def __init__(self, model: str = "claude-sonnet-4-6", api_key: str = None):
        import anthropic
        self._model = model
        self._client = anthropic.Anthropic(
            api_key=api_key or os.environ.get("ANTHROPIC_API_KEY", ""),
        )

    def complete(self, system, user):
        t0 = time.perf_counter()
        response = self._client.messages.create(
            model=self._model,
            max_tokens=300,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        duration = time.perf_counter() - t0
        log.info("  [timing] API: %.2fs", duration)
        content = response.content[0].text.strip()
        return content, duration

    @property
    def display_name(self):
        return f"Anthropic ({self._model})"


# ── Custom / Self-hosted (OpenAI-compatible) ───────────────────────────
class CustomOpenAIProvider(LLMProvider):
    """Any OpenAI-compatible API (vLLM, Ollama, LiteLLM, etc.)."""

    def __init__(self, base_url: str, model: str, api_key: str = "not-needed"):
        from openai import OpenAI
        self._model = model
        self._client = OpenAI(base_url=base_url, api_key=api_key)

    def complete(self, system, user):
        chunks = []
        t0 = time.perf_counter()
        ttft = None
        stream = self._client.chat.completions.create(
            model=self._model,
            max_tokens=300,
            stream=True,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        for chunk in stream:
            if chunk.choices and chunk.choices[0].delta.content:
                if ttft is None:
                    ttft = time.perf_counter() - t0
                chunks.append(chunk.choices[0].delta.content)
        duration = time.perf_counter() - t0
        log.info("  [timing] API: %.2fs | first token: %.2fs",
                 duration, ttft or duration)
        return "".join(chunks).strip(), duration

    @property
    def display_name(self):
        return f"Custom ({self._model})"


# ── Factory ────────────────────────────────────────────────────────────
PROVIDER_MAP = {
    "azure_openai": AzureOpenAIProvider,
    "openai": OpenAIProvider,
    "anthropic": AnthropicProvider,
    "custom": CustomOpenAIProvider,
}


def create_provider(config: dict) -> LLMProvider:
    """Create an LLM provider from a config dict.

    Config examples:
        {"provider": "azure_openai", "endpoint": "https://...", "deployment": "gpt-4.1", "api_key": "..."}
        {"provider": "openai", "model": "gpt-4.1", "api_key": "sk-..."}
        {"provider": "anthropic", "model": "claude-sonnet-4-6", "api_key": "sk-ant-..."}
        {"provider": "custom", "base_url": "http://localhost:11434/v1", "model": "llama3", "api_key": "not-needed"}
    """
    provider_type = config.pop("provider")
    if provider_type not in PROVIDER_MAP:
        raise ValueError(
            f"Unknown provider '{provider_type}'. "
            f"Available: {', '.join(PROVIDER_MAP.keys())}"
        )

    # Resolve env var references in values (e.g. "$AZURE_OPENAI_API_KEY")
    resolved = {}
    for key, val in config.items():
        if isinstance(val, str) and val.startswith("$"):
            resolved[key] = os.environ.get(val[1:], "")
        else:
            resolved[key] = val

    return PROVIDER_MAP[provider_type](**resolved)


def load_provider_from_config(config_path: str = None) -> LLMProvider:
    """Load provider from config file, env vars, or defaults."""
    if config_path is None:
        if getattr(sys, "frozen", False):
            app_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ClipboardCoach")
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(app_dir, "config.json")

    if os.path.exists(config_path):
        with open(config_path) as f:
            config = json.load(f)
        log.info("  Loaded config from %s", config_path)
        return create_provider(config)

    # Fallback: check env vars to auto-detect provider
    if os.environ.get("AZURE_OPENAI_API_KEY"):
        return AzureOpenAIProvider(
            endpoint=os.environ.get("AZURE_OPENAI_ENDPOINT",
                                    "https://foundary-poc-gygiuj.cognitiveservices.azure.com/"),
            deployment=os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4.1"),
            api_key=os.environ["AZURE_OPENAI_API_KEY"],
            api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2025-01-01-preview"),
        )

    if os.environ.get("OPENAI_API_KEY"):
        return OpenAIProvider(
            model=os.environ.get("OPENAI_MODEL", "gpt-4.1"),
        )

    if os.environ.get("ANTHROPIC_API_KEY"):
        return AnthropicProvider(
            model=os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6"),
        )

    raise RuntimeError(
        "No LLM provider configured. Either:\n"
        "  1. Create config.json (see config.example.json)\n"
        "  2. Set AZURE_OPENAI_API_KEY, OPENAI_API_KEY, or ANTHROPIC_API_KEY env var"
    )
