"""Integration tests for ClipFix."""
import json
import os
import sys
import time

# ── Setup ───────────────────────────────────────────────────────────────
API_KEY = os.environ.get("AZURE_OPENAI_API_KEY", "")
if not API_KEY:
    print("FAIL: AZURE_OPENAI_API_KEY not set")
    sys.exit(1)

# Initialize the provider before running tests
from providers import load_provider_from_config
import clipboard_coach
clipboard_coach.provider = load_provider_from_config()

passed = 0
failed = 0


def test(name, fn):
    global passed, failed
    try:
        t0 = time.perf_counter()
        fn()
        elapsed = time.perf_counter() - t0
        print(f"  PASS  {name} ({elapsed:.2f}s)")
        passed += 1
    except AssertionError as e:
        print(f"  FAIL  {name}: {e}")
        failed += 1
    except Exception as e:
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        failed += 1


# ── Tests ───────────────────────────────────────────────────────────────
print("=" * 60)
print("  ClipFix -- Integration Tests")
print("=" * 60)


# 1. Azure OpenAI connectivity
def test_api_connectivity():
    from openai import AzureOpenAI
    client = AzureOpenAI(
        azure_endpoint="https://foundary-poc-gygiuj.cognitiveservices.azure.com/",
        api_key=API_KEY,
        api_version="2025-01-01-preview",
    )
    r = client.chat.completions.create(
        model="gpt-4.1",
        max_tokens=100,
        temperature=0,
        messages=[{"role": "user", "content": "Say OK"}],
    )
    content = r.choices[0].message.content
    assert content and len(content) > 0, f"Empty response: {r.model_dump_json()}"

test("API connectivity", test_api_connectivity)


# 2. Coach analysis returns valid JSON
def test_coach_analysis():
    from clipboard_coach import analyze_message
    result, api_time = analyze_message(
        "Hey I was just wondering if maybe we could possibly push the meeting back a bit"
    )
    assert isinstance(result, dict), f"Expected dict, got {type(result)}"
    assert result["verdict"] in ("improve", "good"), f"Bad verdict: {result['verdict']}"
    assert "nudge" in result, "Missing nudge"
    assert api_time > 0, "API time should be positive"

test("Coach analysis (improve case)", test_coach_analysis)


# 3. Good message detection
def test_good_message():
    from clipboard_coach import analyze_message
    result, _ = analyze_message(
        "The deployment is scheduled for Friday at 3 PM. Please confirm your availability."
    )
    assert isinstance(result, dict), f"Expected dict, got {type(result)}"
    assert result["verdict"] in ("improve", "good"), f"Bad verdict: {result['verdict']}"

test("Coach analysis (good case)", test_good_message)


# 4. Message filter -- should detect
def test_filter_detects_messages():
    from clipboard_coach import looks_like_message
    messages = [
        "Hey team, I think we should revisit the project timeline",
        "Hi, just wanted to follow up on my last email about the budget",
        "Could you please review the attached document and let me know your thoughts",
    ]
    for msg in messages:
        assert looks_like_message(msg), f"Should detect as message: {msg[:50]}"

test("Filter detects natural messages", test_filter_detects_messages)


# 5. Message filter -- should reject
def test_filter_rejects_non_messages():
    from clipboard_coach import looks_like_message
    non_messages = [
        "import os; from pathlib import Path",
        "https://github.com/user/repo/pull/42",
        "C:\\Users\\ankur\\Documents\\file.txt",
        "hello",
        '{"key": "value", "nested": {"a": 1}}',
        "function getData() { return fetch(url); }",
    ]
    for msg in non_messages:
        assert not looks_like_message(msg), f"Should reject: {msg[:50]}"

test("Filter rejects code/URLs/short text", test_filter_rejects_non_messages)


# 6. Cache works
def test_cache():
    from clipboard_coach import analyze_message, _cache
    _cache.clear()
    msg = "I just wanted to check in and see if everything is going okay with the project"
    result1, t1 = analyze_message(msg)
    result2, t2 = analyze_message(msg)
    assert t2 == 0.0, f"Second call should be cached (0s), got {t2:.2f}s"
    assert result1 == result2, "Cached result should match"

test("Cache returns instant result", test_cache)


# 7. Response time
def test_response_time():
    from clipboard_coach import analyze_message, _cache
    _cache.clear()
    t0 = time.perf_counter()
    result, api_time = analyze_message(
        "Please make sure to complete the review before end of day"
    )
    total = time.perf_counter() - t0
    print(f"         API: {api_time:.2f}s | Total: {total:.2f}s")
    assert total < 15, f"Too slow: {total:.1f}s (should be <15s)"

test("Response time < 15s", test_response_time)


# 8. Notification shows rewrite text
def test_notification_with_rewrite():
    from clipboard_coach import analyze_message, display_result, _cache
    _cache.clear()
    result, _ = analyze_message(
        "Hey I was just wondering if maybe we could possibly push the meeting back a bit"
    )
    assert result["verdict"] == "improve", f"Expected improve, got {result['verdict']}"
    assert result.get("rewrite"), "Expected a rewrite but got none"
    rewrite = result["rewrite"]
    assert len(rewrite) < 200, f"Rewrite too long for inline display: {len(rewrite)} chars"
    # display_result sends the notification -- visually confirm it shows the rewrite
    display_result(result)
    print(f"         Rewrite in notification: {rewrite[:100]}")

test("Notification shows rewrite (check toast)", test_notification_with_rewrite)


# ── Summary ─────────────────────────────────────────────────────────────
print("=" * 60)
print(f"  Results: {passed} passed, {failed} failed")
print("=" * 60)
sys.exit(1 if failed else 0)
