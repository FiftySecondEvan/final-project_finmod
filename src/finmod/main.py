"""Entry point for the FinMod package."""

from __future__ import annotations

import argparse
from pathlib import Path

from .modeler import (
    Assumptions,
    format_table,
    infer_assumptions,
    load_income_statement,
    project_statement,
    write_projections_to_xlsx,
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
            "Infer assumptions from a baseline income statement Excel template "
            "and project future periods."
        )
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=Path("Inputs_Historical/Baseline IS.xlsx"),
        help="Path to the baseline .xlsx template (default: Inputs_Historical/Baseline IS.xlsx).",
    )
    parser.add_argument(
        "--years",
        type=int,
        nargs="*",
        help="Optional subset of years to display; defaults to all columns in the sheet.",
    )
    parser.add_argument(
        "--output-xlsx",
        type=Path,
        default=Path("Outputs_Projections/Projected IS.xlsx"),
        help=(
            "Path to write an XLSX file with assumptions and projections "
            "(default: Outputs_Projections/Projected IS.xlsx)."
        ),
    )
    return parser.parse_args()


def _render_assumptions(assumptions: Assumptions) -> str:
    lines = [
        "AI-inferred assumptions (deterministic):",
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
    if not args.file.exists():
        raise SystemExit(f"File not found: {args.file}")

    income_statement = load_income_statement(args.file)
    assumptions = infer_assumptions(income_statement)
    projected_series = project_statement(income_statement, assumptions)
    years_to_show = args.years or income_statement.years

    print(_render_assumptions(assumptions))
    print("\nProjected income statement:\n")
    print(format_table(projected_series, years_to_show))

    # Ensure output directory exists
    if args.output_xlsx:
        output_path = _next_versioned(args.output_xlsx)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        # Preserve the original layout by writing into a copy of the baseline template
        write_template_with_projections(args.file, output_path, assumptions, projected_series)
        print(f"\nSaved projections to {output_path}")


if __name__ == "__main__":
    run()
