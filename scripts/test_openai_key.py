"""
Quick sanity check to verify OPENAI_API_KEY works.

Loads .env if present, then performs a tiny chat completion call.
Exit code 0 on success, 1 on failure.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from openai import OpenAI


def _load_dotenv(path: Path) -> None:
    """Minimal .env loader to populate os.environ."""
    if not path.exists():
        return
    for line in path.read_text().splitlines():
        if not line or line.strip().startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key, val = key.strip(), val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


def main() -> None:
    _load_dotenv(Path(".env"))
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("OPENAI_API_KEY missing. Set it in .env or env vars.")
        sys.exit(1)

    client = OpenAI(api_key=api_key)
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": "Write a 3-line haiku about data flowing like a river."}],
            max_tokens=40,
            temperature=0,
        )
        text = resp.choices[0].message.content
        print("Success. Model replied:", text)
        sys.exit(0)
    except Exception as exc:
        print("OpenAI call failed:", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
