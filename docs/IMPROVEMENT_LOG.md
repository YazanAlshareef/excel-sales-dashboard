# Improvement log

## 2026-11-15, cycle 1

Baseline: 8 tests pass and the offline build runs clean (raw 8495, clean 8335, log 4545). The Excel stage was not run, because it needs desktop Excel. No linter or type checker is configured.

### Done this cycle (branches awaiting review, not merged)

| Item | Branch | Result |
|---|---|---|
| Add a test for the rule counts on the sample data (a5d35d3). Tests only. | test/rule-counts-sample | 9 tests |
| Fix the truncation length in the README (4725c6d). Docs only. Long product names are cut at 28 characters, not 30, checked against the code. | docs/truncation-length | Docs only |
| prepare_data: usage message and row count warning (8077e71). Behaviour change: a clear usage message instead of a traceback. Extra arguments also print the usage message, which the commit message does not mention. | fix/prepare-data-usage | Usage and warning paths not run |

Each branch was reviewed on its own against main (9503ecb).

### Not verified

- The usage and row-count warning paths in prepare_data were not run, because the 23 MB source file is not in the repo.
- Recalculation in Excel, the dropdowns by hand, and other Excel versions or languages.
- Test counts on the docs and prepare_data branches were not recorded for this log.
- Lint and type results.

### Backlog, top 5 not yet done

1. finish_in_excel.py saves the final file even when its error scan finds errors (scripts/finish_in_excel.py:77, 91, 97-98, 101). Needs desktop Excel to check. Effort S, risk low.
2. Top 10 products and top countries show zero-value rows under narrow filters (scripts/build_workbook.py:187, 195-197, 206, 210). Inferred from the formulas, not run in Excel. Effort M, risk medium.
3. Nothing checks the workbook the build writes (.github/workflows/ci.yml:16, scripts/build_workbook.py:25-455). Effort M, risk low.
4. Docs say every KPI is a SUMIFS, but two cards are ratios (WALKTHROUGH.md:56, README.md:87, scripts/build_workbook.py:378-379). Effort S, risk low.
5. README does not say the top countries chart ignores the Country dropdown (scripts/build_workbook.py:202, README.md:141-143). Effort S, risk low.

### Lessons

- The README said 30 and the formula cuts at 28 (README.md:145, scripts/build_workbook.py:196). Numbers in docs are worth checking against the code.
- The prepare_data commit message missed one behaviour, so a commit message should list every behaviour change.
- The sample file is not in the repo, so two paths stay untested here.

### Diagrams

Diagrams live on the docs/diagrams branch, in docs/diagrams/ (build-flow.md, cleaning-rules.md, components.md, dashboard-filter.md, workbook-sheets.md).

### Status

All branches are local and await the owner's approval. Nothing is pushed or merged.
