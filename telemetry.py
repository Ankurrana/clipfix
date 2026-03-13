"""Telemetry module for ClipFix usage tracking.

Stores structured events (session starts, analyses, rewrite pastes) in a
JSONL file for later study and improvement.
"""

import json
import logging
import threading
from datetime import datetime, timedelta
from pathlib import Path

log = logging.getLogger("coach")

# ── Event Types ──────────────────────────────────────────────────────────
EVENT_SESSION_START = "session_start"
EVENT_ANALYSIS = "analysis"
EVENT_REWRITE_PASTED = "rewrite_pasted"


class Telemetry:
    """Append-only JSONL telemetry writer (thread-safe)."""

    def __init__(self, app_dir: Path):
        self._file = app_dir / "telemetry.jsonl"
        self._lock = threading.Lock()
        self._session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self._analysis_count = 0

    # ── internal ─────────────────────────────────────────────────────────
    def _append(self, event: dict):
        event["session_id"] = self._session_id
        event["timestamp"] = datetime.now().isoformat()
        with self._lock:
            try:
                with open(self._file, "a", encoding="utf-8") as f:
                    f.write(json.dumps(event, ensure_ascii=False) + "\n")
            except Exception as e:
                log.warning("  [telemetry] write failed: %s", e)

    # ── public API ───────────────────────────────────────────────────────
    def log_session_start(self, provider_name: str):
        """Record that the app was launched."""
        self._append({
            "event": EVENT_SESSION_START,
            "provider": provider_name,
        })

    def log_analysis(self, input_text: str, result: dict,
                     api_duration: float, total_duration: float,
                     cached: bool):
        """Record a full analysis cycle (input + LLM output)."""
        self._analysis_count += 1
        self._append({
            "event": EVENT_ANALYSIS,
            "analysis_number": self._analysis_count,
            "input_text": input_text,
            "verdict": result.get("verdict"),
            "issue": result.get("issue"),
            "nudge": result.get("nudge"),
            "rewrite": result.get("rewrite"),
            "api_duration_s": round(api_duration, 3),
            "total_duration_s": round(total_duration, 3),
            "cached": cached,
        })

    def log_rewrite_pasted(self, rewrite_text: str):
        """Record that the user accepted a rewrite via Ctrl+M."""
        self._append({
            "event": EVENT_REWRITE_PASTED,
            "rewrite": rewrite_text,
        })

    # ── stats (for summary / tray menu) ──────────────────────────────────
    def session_analysis_count(self) -> int:
        return self._analysis_count

    def load_all_events(self) -> list[dict]:
        """Read all telemetry events from disk."""
        if not self._file.exists():
            return []
        events = []
        with open(self._file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        events.append(json.loads(line))
                    except json.JSONDecodeError:
                        continue
        return events

    def _filter_since(self, events: list[dict], since: datetime) -> list[dict]:
        """Filter events to those after a given datetime."""
        cutoff = since.isoformat()
        return [e for e in events if e.get("timestamp", "") >= cutoff]

    def _compute_stats(self, events: list[dict]) -> dict:
        """Compute stats from a list of events."""
        analyses = [e for e in events if e["event"] == EVENT_ANALYSIS]
        pastes = [e for e in events if e["event"] == EVENT_REWRITE_PASTED]
        sessions = [e for e in events if e["event"] == EVENT_SESSION_START]

        issue_counts: dict[str, int] = {}
        verdicts = {"good": 0, "improve": 0}
        for a in analyses:
            v = a.get("verdict", "")
            if v in verdicts:
                verdicts[v] += 1
            issue = a.get("issue")
            if issue:
                issue_counts[issue] = issue_counts.get(issue, 0) + 1

        top_issues = sorted(issue_counts.items(), key=lambda x: -x[1])
        total = verdicts["good"] + verdicts["improve"]

        return {
            "total_sessions": len(sessions),
            "total_analyses": len(analyses),
            "total_rewrites_pasted": len(pastes),
            "verdicts": verdicts,
            "top_issues": top_issues[:10],
            "clean_rate": round(verdicts["good"] / total * 100) if total else 0,
            "acceptance_rate": (
                round(len(pastes) / verdicts["improve"] * 100, 1)
                if verdicts["improve"] > 0 else 0
            ),
        }

    def summary(self) -> dict:
        """Compute aggregate stats from all telemetry data."""
        return self._compute_stats(self.load_all_events())

    def weekly_stats(self) -> dict:
        """Stats for the last 7 days."""
        events = self.load_all_events()
        since = datetime.now() - timedelta(days=7)
        return self._compute_stats(self._filter_since(events, since))

    def prev_weekly_stats(self) -> dict:
        """Stats for 8-14 days ago (prior week for comparison)."""
        events = self.load_all_events()
        now = datetime.now()
        start = now - timedelta(days=14)
        end = now - timedelta(days=7)
        window = [
            e for e in self._filter_since(events, start)
            if e.get("timestamp", "") < end.isoformat()
        ]
        return self._compute_stats(window)

    def startup_summary(self) -> str | None:
        """One-line startup summary from last 7 days. None if no data."""
        stats = self.weekly_stats()
        if stats["total_analyses"] == 0:
            return None

        parts = [f"{stats['total_analyses']} messages analyzed"]
        parts.append(f"{stats['clean_rate']}% clean")

        if stats["top_issues"]:
            top_issue, top_count = stats["top_issues"][0]
            parts.append(f"top issue: {top_issue} ({top_count}x)")

        # Trend vs prior week
        prev = self.prev_weekly_stats()
        if prev["total_analyses"] > 0:
            delta = stats["clean_rate"] - prev["clean_rate"]
            if delta > 0:
                parts.append(f"up {delta}pp from last week")
            elif delta < 0:
                parts.append(f"down {abs(delta)}pp from last week")

        return "Last 7 days: " + ", ".join(parts)

    def weekly_digest(self) -> str | None:
        """Multi-line weekly digest. None if no data."""
        stats = self.weekly_stats()
        if stats["total_analyses"] == 0:
            return None

        prev = self.prev_weekly_stats()
        lines = [f"This week: {stats['total_analyses']} messages"]

        total = stats["verdicts"]["good"] + stats["verdicts"]["improve"]
        lines.append(
            f"  {stats['verdicts']['good']}/{total} clean ({stats['clean_rate']}%)"
        )

        if stats["total_rewrites_pasted"] > 0:
            lines.append(
                f"  {stats['total_rewrites_pasted']} rewrites applied "
                f"({stats['acceptance_rate']}% acceptance)"
            )

        # Trend comparison
        if prev["total_analyses"] > 0:
            delta = stats["clean_rate"] - prev["clean_rate"]
            if delta > 0:
                lines.append(f"  Trend: {delta}pp improvement from last week!")
            elif delta < 0:
                lines.append(f"  Trend: {abs(delta)}pp dip from last week")
            else:
                lines.append("  Trend: holding steady")

        # Top issues this week
        if stats["top_issues"]:
            lines.append("  Top issues:")
            for issue, count in stats["top_issues"][:3]:
                lines.append(f"    {issue}: {count}x")

        return "\n".join(lines)

    def should_show_weekly_digest(self) -> bool:
        """True if it's been 7+ days since the last digest was shown."""
        marker = self._file.parent / ".last_weekly_digest"
        if not marker.exists():
            return self.weekly_stats()["total_analyses"] > 0
        try:
            last = datetime.fromisoformat(marker.read_text().strip())
            return datetime.now() - last >= timedelta(days=7)
        except Exception:
            return True

    def mark_weekly_digest_shown(self):
        """Record that the weekly digest was just shown."""
        marker = self._file.parent / ".last_weekly_digest"
        marker.write_text(datetime.now().isoformat())
