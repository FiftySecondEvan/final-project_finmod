"""
AI-assisted CLI: uses OpenAI to propose assumptions and projections.

Falls back to deterministic assumptions if the API call fails or parsing fails.
"""

from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import asdict
from pathlib import Path
from typing import Dict, Optional

from openai import OpenAI

from finmod.modeler import Assumptions, extract_all_series, format_table, infer_dynamic_assumptions, project_dynamic, write_template_with_projections


def _next_versioned(path: Path) -> Path:
    """Return path, or a numbered variant if it already exists."""
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    n = 1
    while True:
        candidate = path.with_name(f"{stem} v{n}{suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Use OpenAI to infer assumptions from a baseline IS template and project future periods. "
            "Requires OPENAI_API_KEY; falls back to deterministic mode if the call fails."
        )
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=Path("Inputs_Historical/Baseline IS.xlsx"),
        help="Path to the baseline .xlsx template (default: Inputs_Historical/Baseline IS.xlsx).",
    )
    parser.add_argument(
        "--output-xlsx",
        type=Path,
        default=Path("Outputs_Projections/Projected IS (AI).xlsx"),
        help=(
            "Path to write an XLSX file with assumptions and projections "
            "(default: Outputs_Projections/Projected IS (AI).xlsx)."
        ),
    )
    parser.add_argument(
        "--model",
        default="gpt-4o-mini",
        help="OpenAI model name (default: gpt-4o-mini).",
    )
    return parser.parse_args()


def _load_dotenv(path: Path) -> None:
    """Minimal .env loader to populate os.environ without extra deps."""
    if not path.exists():
        return
    for line in path.read_text().splitlines():
        if not line or line.strip().startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key, val = key.strip(), val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


def _build_prompt(years, series_map) -> str:
    """Create a compact text prompt summarizing historicals and asking for JSON assumptions."""
    lines = []
    lines.append("You are a financial analyst. Given historical P&L lines by year, infer forward assumptions.")
    lines.append("Return ONLY JSON with keys: revenue_growth_cagr (decimal), ratios (object mapping line label to decimal pct of revenue), and notes (string).")
    lines.append("All values must be decimals (e.g., 0.12 for 12%).")
    lines.append("Historicals:")
    for label, series in series_map.items():
        parts = [f"{y}:{series.get(y)}" for y in years if y in series]
        if parts:
            lines.append(f"{label}: " + ", ".join(parts))
    lines.append(
        "Assume steady-state growth/margins aligned with recent performance. If a label seems non-operating (e.g., Interest Expense, Tax), still return a pct of revenue as an approximation."
    )
    return "\n".join(lines)


def _call_openai_for_assumptions(years, series_map, model: str) -> Dict:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set.")

    client = OpenAI(api_key=api_key)
    prompt = _build_prompt(years, series_map)
    resp = client.chat.completions.create(
        model=model,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You are a precise financial analyst. Be concise and return JSON only."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )
    raw_content = resp.choices[0].message.content

    # openai>=1.0 may return content as a list of parts; normalize to string
    if isinstance(raw_content, list):
        content = "".join(part.get("text", "") if isinstance(part, dict) else str(part) for part in raw_content)
    else:
        content = str(raw_content) if raw_content is not None else ""

    try:
        data: Dict[str, float] = json.loads(content)
    except Exception as exc:
        # try to extract first JSON object from the string as a fallback
        match = re.search(r"{.*}", content, flags=re.DOTALL)
        if match:
            data = json.loads(match.group(0))
        else:
            raise RuntimeError(f"Failed to parse JSON from OpenAI response: {content}") from exc
    return data


def _render_assumptions(source: str, model_name: str, revenue_cagr: float, ratios: Dict[str, float]) -> str:
    lines = [f"AI-inferred assumptions (source: {source}, model: {model_name}):"]
    lines.append(f"- Revenue CAGR: {revenue_cagr*100:.2f}%")
    for label, pct in list(ratios.items())[:8]:
        lines.append(f"- {label}: {pct*100:.2f}% of revenue")
    return "\n".join(lines)


def run() -> None:
    args = _parse_args()

    # Load .env from project root if present
    _load_dotenv(Path(".env"))

    if not args.file.exists():
        raise SystemExit(f"File not found: {args.file}")

    years, revenue_label, series_map = extract_all_series(args.file)

    # Deterministic baseline assumptions for all lines
    det_revenue_cagr, det_strategies = infer_dynamic_assumptions(years, revenue_label, series_map)

    # Try OpenAI, fall back to deterministic inference
    source = "openai"
    model_used = args.model
    ai_ratios: Dict[str, float] = {}
    note = ""
    ai_revenue_cagr = det_revenue_cagr
    try:
        data = _call_openai_for_assumptions(years, series_map, args.model)
        ai_revenue_cagr = float(data.get("revenue_growth_cagr", det_revenue_cagr))
        ai_ratios = {k: float(v) for k, v in data.get("ratios", {}).items()}
        note = data.get("notes", "")
    except Exception as exc:
        source = f"fallback (deterministic) due to error: {exc}"
        model_used = "n/a"

    # Merge AI ratios into strategies
    strategies = dict(det_strategies)
    for label, pct in ai_ratios.items():
        strategies[label] = ("ratio", pct)

    projected_series = project_dynamic(years, revenue_label, series_map, ai_revenue_cagr, strategies)
    years_to_show = years

    print(_render_assumptions(source, model_used, ai_revenue_cagr, ai_ratios or {revenue_label: 1.0}))
    print("\nProjected income statement:\n")
    print(format_table(projected_series, years_to_show))

    if args.output_xlsx:
        output_path = _next_versioned(args.output_xlsx)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        note_text = f"Generated via finmod_ai using model: {model_used}. {note}".strip()
        # Populate template assumption slots with AI ratios when available
        assumptions_obj = Assumptions(
            revenue_growth_cagr=ai_revenue_cagr,
            cogs_pct=ai_ratios.get("COGS", 0.0),
            sgna_pct=ai_ratios.get("SG&A", 0.0),
            rnd_pct=ai_ratios.get("R&D", 0.0),
            other_income_pct=ai_ratios.get("Other Income", 0.0),
            capex_pct=ai_ratios.get("Capex", ai_ratios.get("CAPEX", 0.0)),
        )
        write_template_with_projections(
            args.file,
            output_path,
            assumptions_obj,
            projected_series,
            note=note_text,
            ratios=ai_ratios,
        )
        print(f"\nSaved projections to {output_path}")


if __name__ == "__main__":
    run()
