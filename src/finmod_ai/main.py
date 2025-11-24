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

from finmod.modeler import (
    Assumptions,
    format_table,
    infer_assumptions,
    load_income_statement,
    project_statement,
    write_template_with_projections,
)


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


def _build_prompt(income_statement) -> str:
    """Create a compact text prompt summarizing historicals and asking for JSON assumptions."""
    lines = []
    lines.append("You are a financial analyst. Given historical P&L lines by year, infer forward assumptions.")
    lines.append("Return ONLY JSON with keys: revenue_growth_cagr, cogs_pct, sgna_pct, rnd_pct, other_income_pct, capex_pct.")
    lines.append("All values must be decimals (e.g., 0.12 for 12%).")
    lines.append("Historicals:")
    years = sorted(income_statement.revenue)
    for y in years:
        lines.append(
            f"{y}: revenue={income_statement.revenue.get(y)}, "
            f"cogs={income_statement.cogs.get(y)}, "
            f"sgna={income_statement.sgna.get(y)}, "
            f"rnd={income_statement.rnd.get(y)}, "
            f"other_income={income_statement.other_income.get(y)}, "
            f"capex={income_statement.capex.get(y)}"
        )
    lines.append("Assume steady-state growth/margins aligned with recent performance.")
    return "\n".join(lines)


def _call_openai_for_assumptions(income_statement, model: str) -> Assumptions:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set.")

    client = OpenAI(api_key=api_key)
    prompt = _build_prompt(income_statement)
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

    return Assumptions(
        revenue_growth_cagr=float(data["revenue_growth_cagr"]),
        cogs_pct=float(data["cogs_pct"]),
        sgna_pct=float(data["sgna_pct"]),
        rnd_pct=float(data["rnd_pct"]),
        other_income_pct=float(data["other_income_pct"]),
        capex_pct=float(data["capex_pct"]),
    )


def _render_assumptions(assumptions: Assumptions, source: str, model_name: str) -> str:
    lines = [
        f"AI-inferred assumptions (source: {source}, model: {model_name}):",
        f"- Revenue CAGR: {assumptions.revenue_growth_cagr*100:.2f}%",
        f"- COGS: {assumptions.cogs_pct*100:.2f}% of revenue",
        f"- SG&A: {assumptions.sgna_pct*100:.2f}% of revenue",
        f"- R&D: {assumptions.rnd_pct*100:.2f}% of revenue",
        f"- Other income: {assumptions.other_income_pct*100:.2f}% of revenue",
        f"- Capex: {assumptions.capex_pct*100:.2f}% of revenue",
    ]
    return "\n".join(lines)


def run() -> None:
    args = _parse_args()

    # Load .env from project root if present
    _load_dotenv(Path(".env"))

    if not args.file.exists():
        raise SystemExit(f"File not found: {args.file}")

    income_statement = load_income_statement(args.file)

    # Try OpenAI, fall back to deterministic inference
    source = "openai"
    model_used = args.model
    try:
        assumptions = _call_openai_for_assumptions(income_statement, args.model)
    except Exception as exc:
        source = f"fallback (deterministic) due to error: {exc}"
        model_used = "n/a"
        assumptions = infer_assumptions(income_statement)

    projected_series = project_statement(income_statement, assumptions)
    years_to_show = income_statement.years

    print(_render_assumptions(assumptions, source, model_used))
    print("\nProjected income statement:\n")
    print(format_table(projected_series, years_to_show))

    if args.output_xlsx:
        output_path = _next_versioned(args.output_xlsx)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        note = f"Generated via finmod_ai using model: {model_used}"
        write_template_with_projections(args.file, output_path, assumptions, projected_series, note=note)
        print(f"\nSaved projections to {output_path}")


if __name__ == "__main__":
    run()
