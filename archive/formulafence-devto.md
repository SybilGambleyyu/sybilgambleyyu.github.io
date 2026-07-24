# Spreadsheet changes that can pass CI

For important spreadsheets, the painful question is rarely “which file changed?” It is “did someone replace the calculation behind this number, what now depends on it, and did we notice before the model was shared?”

Excel is flexible because the model and its logic live together. That same flexibility turns ordinary version history into a weak review surface: Git sees a binary file, while a reviewer needs to see a change in intent.

[Workbook-aware Git diffs](https://www.xltrail.com/git-xl) and [formula anomaly analysis](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/) already address valuable parts of the problem. I wanted the missing merge-boundary layer: a small, local tool that could describe a semantic workbook change, trace its visible consequences, and enforce a rule that a reviewer can read.

That tool is [FormulaFence](https://github.com/SybilGambleyyu/formulafence), an MIT-licensed, local-first CLI for `.xlsx` and `.xlsm` workbooks.

## Put the control at the merge boundary

FormulaFence compares two workbooks without executing formulas or macros. It detects formula-to-value substitutions, formula changes, sheet visibility and defined-name changes, explicit external references, broken `#REF!` formulas, calculation-setting changes, and macro payload changes.

For every changed cell, it follows statically visible A1-style dependencies and reports downstream formula cells with deterministic shortest-path samples.

```bash
formulafence check approved.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format sarif --output results.sarif
```

The policy is plain YAML. It can protect an output cell, limit routine edits to an input range, ban formula overrides, and cap downstream impact. It does not replace judgment; it makes a review expectation visible and repeatable.

```yaml
version: 1
rules:
  no_formula_to_value: true
  no_new_external_links: true
  no_new_broken_references: true
  no_new_parser_warnings: true
  max_downstream_impact: 100

protected_cells:
  - Dashboard!B12
```

## Fail closed when the analysis has a blind spot

Static analysis has limits. `INDIRECT`, named formulas, structured table references, add-ins, and other Excel features can conceal dependencies. FormulaFence does not invent a graph it cannot justify.

One practical safeguard is parser coverage. When the workbook parser encounters an OOXML extension it cannot fully interpret, FormulaFence records a coverage note. A candidate that adds one can be rejected with `no_new_parser_warnings`.

## Test beyond toy files

Unit fixtures are necessary, but an Office-file reader also needs to meet real workbooks. I validated FormulaFence against the public [Foresight Cap Table and Exit Waterfall Tool](https://github.com/foresighthq/cap-table-tool): 18 sheets, 6,623 non-empty cells, and 4,228 formula cells.

The inspection found an unsupported OOXML extension and recorded it as a structured coverage note instead of leaking a raw dependency warning into CI output. On a local, non-distributed copy, replacing one exit-waterfall formula with a number traced 330 downstream formula cells; the starter policy rejected both the formula override and the impact limit.

Those results are a compatibility demonstration, not a claim that the source model is correct. The full limits and validation record are in the [FormulaFence repository](https://github.com/SybilGambleyyu/formulafence/blob/main/docs/validation.md).

## What it does not claim

FormulaFence does not calculate Excel or prove a financial model correct. Material models still need a qualified owner, recalculation in the approved spreadsheet engine, and independent review.

But a review process should at least make it hard to silently replace a formula with a number. That is the narrow, useful boundary FormulaFence is built to enforce.

The current release is [available on GitHub](https://github.com/SybilGambleyyu/formulafence/releases/tag/v0.2.0). The canonical version of this post lives at [sybilgambleyyu.github.io/posts/formulafence.html](https://sybilgambleyyu.github.io/posts/formulafence.html).
