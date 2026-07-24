# Spreadsheet changes that can pass CI

For important spreadsheets, the painful question is rarely “which file changed?” It is “did someone replace the calculation behind this number, what now depends on it, and did we notice before the model was shared?”

Excel is flexible because the model and its logic live together. That same flexibility turns ordinary version history into a weak review surface: Git sees a binary file, while a reviewer needs to see a change in intent.

[Workbook-aware Git diffs](https://www.xltrail.com/git-xl) and [formula anomaly analysis](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/) already address valuable parts of the problem. I wanted the missing merge-boundary layer: a small, local tool that could describe a semantic workbook change, trace its visible consequences, and enforce a rule that a reviewer can read.

That tool is [FormulaFence](https://github.com/SybilGambleyyu/formulafence), an MIT-licensed, local-first CLI for `.xlsx` and `.xlsm` workbooks.

## Put the control at the merge boundary

FormulaFence compares two workbooks without executing formulas or macros. It detects formula-to-value substitutions, formula changes, sheet visibility, defined-name and Excel-table definition changes, explicit external references, broken `#REF!` formulas, calculation-setting changes, and macro payload changes.

For every changed cell, it follows statically visible A1-style, ordinary named-range, safely expandable formula-defined-name and named `LAMBDA`, `LET`/inline-`LAMBDA`, supported table, and direct 3-D worksheet dependencies and reports downstream formula cells with deterministic shortest-path samples.

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
  no_new_unresolved_references: true
  no_new_dynamic_references: true
  no_table_definition_changes: true
  no_3d_reference_scope_changes: true
  max_downstream_impact: 100

protected_cells:
  - Dashboard!B12
```

## Fail closed when the analysis has a blind spot

FormulaFence 0.9.0 resolves ordinary workbook and sheet-local names with static A1 destinations, plus a conservative set of Excel-table references: table names, static columns or contiguous column ranges, and `#All`/`#Data`/`#Headers`/`#Totals` regions. Profiles inventory the table definitions behind those references, and a diff emits `FF013` if one changes; `no_table_definition_changes` can make that a hard boundary.

It also expands a direct internal 3-D A1 reference such as `Jan:Mar!B2` across every worksheet tab between its endpoints. Profiles identify those formulas. Since inserting, moving, or removing a tab can change that span without changing the formula text, a diff emits `FF014`; `no_3d_reference_scope_changes` can make that a hard boundary. This follows [Excel's documented 3-D-reference behavior](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets). External, malformed, endpoint-missing, and non-A1 variants remain explicit coverage limits rather than guessed dependencies.

Version 0.7.0 follows the static paths behind formula-defined names as well. A name such as `DiscountedValue` can itself use a second name such as `TaxRate`; when the complete definition resolves to internal dependencies or a constant, FormulaFence traces through to the source cells. It handles nested workbook and sheet-local definitions without evaluating the formulas. Relative, cyclic, external, dynamic, 3-D-in-name, and tokenizer-unsupported definitions remain explicit coverage gaps instead of guessed paths. That matches [Excel's support for names representing formulas and constants](https://support.microsoft.com/en-us/excel/names-in-formulas).

Version 0.8.0 removes a common modern false gap: `LET` names and inline `LAMBDA` parameters are local variables, not workbook names. FormulaFence follows their lexical scope—including nested lambdas supplied to functions such as `REDUCE`—and retains the actual A1, named, and table inputs around them. It follows Excel's documented [LET scope](https://support.microsoft.com/en-us/excel/functions/let-function) and [LAMBDA parameter syntax](https://support.microsoft.com/en-us/excel/functions/lambda-function), without evaluating either.

Version 0.9.0 follows a named `LAMBDA` only when its whole definition is static and internal. A caller such as `=ToCelsius(A2)` retains its `A2` edge and gains the function body's static inputs; nested named functions and formula-defined names that call one work the same way. It respects workbook and worksheet-local scope, and recognizes both ordinary Name Manager text and the `_xlfn.LAMBDA`/`_xlpm.`/`_xlop.` OOXML spelling emitted by Excel-compatible writers. That is the same named-function model described in [Excel's LAMBDA documentation](https://support.microsoft.com/en-us/excel/functions/lambda-function). Recursive, dynamic, relative, external, 3-D, and tokenizer-unsupported definitions remain explicit coverage notes rather than guessed graph edges.

It now traces common row-scoped forms without turning a row calculation into a dependency on every row of a table. Inside a table data cell, `[@[Sales Amount]]` and `[Sales Amount]` bind to that row. Qualified forms such as `Sales[@Amount]` and `Sales[[#This Row],[Amount]:[Rate]]` bind to the named table's data row even from an adjacent cell. That follows [Excel's documented structured-reference semantics](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables); header, total, cross-sheet, ambiguous, and complex bracket-escape cases remain coverage notes instead of guessed dependency paths.

One practical safeguard is coverage visibility. When the workbook parser encounters an OOXML extension it cannot fully interpret, FormulaFence records a coverage note. A candidate that adds one can be rejected with `no_new_parser_warnings`. Profiles also list unresolved range tokens and dynamic reference functions; a change can be rejected with `no_new_unresolved_references` or `no_new_dynamic_references`.

## Test beyond toy files

Unit fixtures are necessary, but an Office-file reader also needs to meet real workbooks. I validated FormulaFence against the public [Foresight Cap Table and Exit Waterfall Tool](https://github.com/foresighthq/cap-table-tool): 18 sheets, 6,623 non-empty cells, and 4,228 formula cells.

The inspection found an unsupported OOXML extension and recorded it as a structured coverage note instead of leaking a raw dependency warning into CI output. It also identified 36 cells using `INDIRECT`, making the model’s dynamic-reference surface explicit. On a local, non-distributed copy, replacing one exit-waterfall formula with a number traced 330 downstream formula cells; the starter policy rejected both the formula override and the impact limit.

I also checked a public structured-reference workbook: FormulaFence identified its one table and 15 table-reference formulas with no unresolved tokens. Changing one table data cell on a local copy traced 16 downstream formulas, including a table total and an output outside the table. For the new row-scoped logic, a controlled workbook changed one input and traced exactly its calculated-row cell, a neighboring qualified row-reference cell, and an external summary—without marking the other two table rows. The [validation record](https://github.com/SybilGambleyyu/formulafence/blob/main/docs/validation.md) describes the boundary and evidence.

Those results are a compatibility demonstration, not a claim that the source model is correct. The full limits and validation record are in the [FormulaFence repository](https://github.com/SybilGambleyyu/formulafence/blob/main/docs/validation.md).

## What it does not claim

FormulaFence does not calculate Excel or prove a financial model correct. Material models still need a qualified owner, recalculation in the approved spreadsheet engine, and independent review. Relative, cyclic, external, and tokenizer-unsupported name definitions, unsupported or ambiguous table syntax, spilled ranges, non-static named LAMBDAs, arbitrary custom functions, and features such as `INDIRECT` remain coverage limits rather than guessed graph edges.

But a review process should at least make it hard to silently replace a formula with a number. That is the narrow, useful boundary FormulaFence is built to enforce.

The current release is [FormulaFence 0.9.0 on GitHub](https://github.com/SybilGambleyyu/formulafence/releases/tag/v0.9.0). The canonical version of this post lives at [sybilgambleyyu.github.io/posts/formulafence.html](https://sybilgambleyyu.github.io/posts/formulafence.html).
