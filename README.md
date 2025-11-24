# FinMod Project

This repository starts as a minimal Python skeleton for financial modeling experiments and utilities.

## Structure
- `src/finmod/`: package code
- `tests/`: lightweight sanity checks
- `requirements.txt`: add dependencies here as needed

## Getting started
1) (Optional) create a virtual environment, e.g. `python3 -m venv .venv && source .venv/bin/activate`
2) Install dependencies: `pip install -r requirements.txt`
3) Run the placeholder module: `python -m finmod`
4) Run tests: `python -m pytest` (install `pytest` first, e.g. `pip install pytest`)

## Using the financial model helper
The CLI reads the provided `Baseline IS.xlsx`, infers simple growth/margin assumptions from the historical columns, and projects the remaining years:

```
python -m finmod --file "Baseline IS.xlsx"
```

You can optionally limit the displayed columns with `--years 2023 2024 2025 2026`.

## Next steps
- Expand the heuristics in `src/finmod/modeler.py` with your preferred logic.
- Add more tests under `tests/` to cover specific modeling rules.
