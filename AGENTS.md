# Repository Guidelines

## Project Structure & Module Organization
- `formd.ipynb` contains the main Form-D preparation notebook; keep exploratory cells grouped by workflow (data cleanup, validation, export).
- `BFYGA1P1_QQ4V.XLS` and `K1 Import Template.xls` live at the repository root as source and template datasets; treat them as read-only inputs and export derivatives to `outputs/` (create if absent).

## Build, Test, and Development Commands
- `jupyter lab` launches an interactive environment; open `formd.ipynb` to iterate on Form-D data preparation.
- `jupyter nbconvert --to notebook --execute formd.ipynb --output outputs/formd_run.ipynb` runs the notebook end to end and saves an executed copy for review.

## Coding Style & Naming Conventions
- Keep notebook cells concise; prefer helper functions with snake_case naming in dedicated code cells.
- Use pandas idioms for data manipulation and document complex transforms with short Markdown notes directly above the relevant cell.
- Store generated files under `outputs/` using lowercase hyphenated names that encode the run context (e.g., `outputs/formd-2024q1.xlsx`).

## Testing Guidelines
- Add lightweight verification cells that assert expected row counts, schema conformity, and critical field completeness.
- When automating runs, include guard checks (e.g., `assert df["issuer_cik"].notna().all()`) so nbconvert halts on data regressions.

## Commit & Pull Request Guidelines
- Write commits in imperative mood (`Add validation checks for issuer CIK`); group related notebook and data changes together.
- For pull requests, provide: purpose summary, key data changes, execution command (with parameters), and attach rendered artifacts from `outputs/` for reviewers.
- Screenshot or describe any filing-ready outputs so reviewers can validate formatting without opening local tools.
