# ClipFix

Always-on clipboard text fixer for Windows. Monitors your clipboard, auto-corrects grammar, punctuation, and spelling, and improves communication impact — all powered by your choice of LLM.

Copy a message before sending it. ClipFix analyzes it instantly, shows a silent Windows notification with feedback, and lets you paste the improved version with **Ctrl+M**.

## How It Works

```
Copy text (Ctrl+C)  →  ClipFix analyzes it  →  Toast notification with feedback  →  Ctrl+M to paste rewrite
```

1. **You copy text** — ClipFix detects it instantly via a Windows clipboard listener (no polling).
2. **LLM analyzes it** — Checks grammar, spelling, punctuation, and communication impact (hedging, passive voice, buried asks, etc.).
3. **Notification appears** — A silent toast shows what's weak and the suggested rewrite.
4. **Ctrl+M to paste** — Press Ctrl+M in any window to paste the improved version directly.

If your message is already good, you'll see a "Looks good — send it!" notification so you know it's safe to send.

## Features

- **Instant clipboard detection** — Event-driven via Windows API, zero CPU when idle
- **Grammar + communication coaching** — Fixes errors and improves impact
- **Formatting preservation** — Keeps numbered lists, bullet points, and line breaks intact
- **HTML clipboard support** — Reads rich text from Teams/Outlook to preserve list formatting
- **Silent notifications** — No annoying sounds, just clean toast popups
- **Rewrite preview** — Short rewrites shown directly in the notification
- **Ctrl+M paste** — Global hotkey pastes the rewrite into any active window
- **Result caching** — Instant results for repeated clipboard content
- **Pattern tracking** — Learns your recurring weak patterns and prioritizes coaching on them
- **Pluggable LLM providers** — Azure OpenAI, OpenAI, Anthropic Claude, or any OpenAI-compatible API
- **Setup wizard** — GUI configuration on first run
- **Standalone installer** — Single .exe, no Python required for end users

## Quick Start

### Option A: Run from source (developers)

```powershell
git clone https://github.com/Ankurrana/clipfix.git
cd clipfix
pip install -r requirements.txt
```

Set your API key and run:

```powershell
$env:AZURE_OPENAI_API_KEY="your-key-here"
python clipboard_coach.py
```

### Option B: Run the standalone .exe (end users)

Download `ClipFix.exe` from the [Releases](https://github.com/Ankurrana/clipfix/releases) page and double-click it. A setup wizard will prompt for your LLM provider and API key on first run.

## Configuration

ClipFix supports multiple LLM providers. Configure via **config.json**, **environment variables**, or the **setup wizard**.

### Option 1: Setup Wizard (easiest)

Just run the app without any configuration. The setup wizard GUI will appear and ask for your provider and API key.

### Option 2: Environment Variables

Set one of these and ClipFix auto-detects the provider:

| Provider | Environment Variable | Optional Variables |
|---|---|---|
| **Azure OpenAI** | `AZURE_OPENAI_API_KEY` | `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_VERSION` |
| **OpenAI** | `OPENAI_API_KEY` | `OPENAI_MODEL` |
| **Anthropic** | `ANTHROPIC_API_KEY` | `ANTHROPIC_MODEL` |

Detection order: Azure OpenAI > OpenAI > Anthropic.

**PowerShell:**
```powershell
$env:AZURE_OPENAI_API_KEY="your-key"
python clipboard_coach.py
```

**Permanent (persists across restarts):**
```cmd
setx AZURE_OPENAI_API_KEY "your-key"
```

### Option 3: config.json

Create a `config.json` file in the same directory as the app. Use `$ENV_VAR` syntax to reference environment variables for secrets.

**Azure OpenAI (Azure AI Foundry):**
```json
{
    "provider": "azure_openai",
    "endpoint": "https://your-resource.cognitiveservices.azure.com/",
    "deployment": "gpt-4.1",
    "api_key": "$AZURE_OPENAI_API_KEY",
    "api_version": "2025-01-01-preview"
}
```

**OpenAI:**
```json
{
    "provider": "openai",
    "model": "gpt-4.1",
    "api_key": "$OPENAI_API_KEY"
}
```

**Anthropic Claude:**
```json
{
    "provider": "anthropic",
    "model": "claude-sonnet-4-6",
    "api_key": "$ANTHROPIC_API_KEY"
}
```

**Custom (Ollama, vLLM, LiteLLM, or any OpenAI-compatible API):**
```json
{
    "provider": "custom",
    "base_url": "http://localhost:11434/v1",
    "model": "llama3",
    "api_key": "not-needed"
}
```

See `config.example.json` for all options.

## Building the Installer

### Prerequisites

- Python 3.10+
- Dependencies: `pip install -r requirements.txt`
- PyInstaller: `pip install pyinstaller`

### Build

```powershell
python build.py
```

This creates `dist/ClipFix.exe` — a standalone 21 MB executable that includes Python and all dependencies.

### Install

Run `install.bat` to:
- Copy the exe to `%LOCALAPPDATA%\ClipFix`
- Create a Start Menu shortcut
- Add auto-start at login (background mode)

Or just run `dist/ClipFix.exe` directly — no installation needed.

### Uninstall

**Option A: Run the uninstaller**

Run `uninstall.bat` — it stops ClipFix, removes the exe, Start Menu shortcut, and auto-start entry.

**Option B: Manual uninstall**

1. Right-click the ClipFix tray icon and click **Quit** (or kill `ClipFix.exe` in Task Manager)
2. Delete the install folder:
   ```powershell
   Remove-Item "$env:LOCALAPPDATA\ClipFix" -Recurse -Force
   ```
3. Remove the Start Menu shortcut:
   ```powershell
   Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\ClipFix.lnk" -Force
   ```
4. Remove the auto-start shortcut:
   ```powershell
   Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\ClipFix.lnk" -Force
   ```

### Rebuild After Code Changes

Run `rebuild.bat` to stop the running app, rebuild, run tests, and reinstall in one step.

## Running as a Background Service

```powershell
python clipboard_coach.py --background
```

In background mode:
- Logs to `clipboard-coach.log` instead of console
- All interaction via notifications and Ctrl+M
- Ideal for auto-start at login

## Running Tests

```powershell
$env:AZURE_OPENAI_API_KEY="your-key"
python test_integration.py
```

Tests cover:
- API connectivity
- Message analysis (improve and good cases)
- Message filter (detects natural language, rejects code/URLs)
- Result caching
- Response time
- Notification with rewrite text

## Architecture

```
clipfix/
  clipboard_coach.py    -- Main app (clipboard listener, notifications, hotkeys)
  providers.py          -- LLM provider interface + implementations
  setup_wizard.py       -- First-run GUI configuration
  config.example.json   -- Example provider configurations
  test_integration.py   -- Integration tests
  build.py              -- PyInstaller build script
  install.bat           -- Windows installer
  uninstall.bat         -- Windows uninstaller
  rebuild.bat           -- Rebuild + reinstall in one step
  installer.iss         -- Inno Setup script (optional, for .exe installer)
```

### Adding a New LLM Provider

1. Create a new class in `providers.py` that extends `LLMProvider`:

```python
class MyProvider(LLMProvider):
    def __init__(self, api_key: str, model: str = "my-model"):
        self._model = model
        # Initialize your client

    def complete(self, system: str, user: str) -> tuple[str, float]:
        # Call your API, return (response_text, duration_seconds)
        ...

    @property
    def display_name(self) -> str:
        return f"MyProvider ({self._model})"
```

2. Register it in `PROVIDER_MAP`:

```python
PROVIDER_MAP = {
    ...
    "my_provider": MyProvider,
}
```

3. Use it in `config.json`:

```json
{
    "provider": "my_provider",
    "model": "my-model",
    "api_key": "$MY_API_KEY"
}
```

## Requirements

- **OS:** Windows 10/11
- **Admin:** Recommended (for global keyboard hooks)
- **LLM:** One of Azure OpenAI, OpenAI, Anthropic, or any OpenAI-compatible API
- **Python:** 3.10+ (only needed for running from source)

## License

MIT
