from pathlib import Path

from finmod.modeler import infer_assumptions, load_income_statement, project_statement


def test_projection_runs_with_baseline_fixture():
    baseline = Path("Inputs_Historical/Baseline IS.xlsx")
    assert baseline.exists(), "Baseline workbook must be present for the demo test."

    income_statement = load_income_statement(baseline)
    assumptions = infer_assumptions(income_statement)
    projected = project_statement(income_statement, assumptions)

    # Validate we produced forecasts for a future year (2026 exists in the template)
    future_year = max(income_statement.years)
    assert future_year in projected["Revenue"]
    assert projected["Revenue"][future_year] > 0
