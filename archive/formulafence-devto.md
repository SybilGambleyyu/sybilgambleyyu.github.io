# Spreadsheet changes that can pass CI

For important spreadsheets, the painful question is rarely “which file changed?” It is “did someone replace the calculation behind this number, what now depends on it, and did we notice before the model was shared?”

Excel is flexible because the model and its logic live together. That same flexibility turns ordinary version history into a weak review surface: Git sees a binary file, while a reviewer needs to see a change in intent.

[Workbook-aware Git diffs](https://www.xltrail.com/git-xl) and [formula anomaly analysis](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/) already address valuable parts of the problem. I wanted the missing merge-boundary layer: a small, local tool that could describe a semantic workbook change, trace its visible consequences, and enforce a rule that a reviewer can read.

That tool is [FormulaFence](https://github.com/SybilGambleyyu/formulafence), an MIT-licensed, local-first CLI for `.xlsx` and `.xlsm` workbooks.

## Put the control at the merge boundary

FormulaFence compares two workbooks without executing formulas or macros. It detects formula-to-value substitutions, formula changes, sheet visibility, defined-name, Excel-table, data-validation, conditional-formatting, and operational protection-control changes, explicit external references, broken `#REF!` formulas, calculation-setting changes, and macro payload changes.

For every changed cell, it follows statically visible A1-style, ordinary named-range, safely expandable formula-defined-name and named `LAMBDA`, direct dynamic-array spill anchors, fixed legacy-CSE outputs, currently observed dynamic-array output members, `LET`/inline-`LAMBDA`, supported table, and direct 3-D worksheet dependencies and reports downstream formula cells with deterministic shortest-path samples.

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
  no_new_spill_references: true
  no_new_dynamic_array_output_references: true
  no_new_implicit_intersections: true
  no_array_formula_semantics_changes: true
  no_new_tokenization_failures: true
  no_table_definition_changes: true
  no_data_validation_changes: true
  no_conditional_formatting_changes: true
  no_protection_changes: true
  no_3d_reference_scope_changes: true
  max_downstream_impact: 100

protected_cells:
  - Dashboard!B12
```

## Fail closed when the analysis has a blind spot

FormulaFence 0.10.0 resolves ordinary workbook and sheet-local names with static A1 destinations, plus a conservative set of Excel-table references: table names, static columns or contiguous column ranges, and `#All`/`#Data`/`#Headers`/`#Totals` regions. Profiles inventory the table definitions behind those references, and a diff emits `FF013` if one changes; `no_table_definition_changes` can make that a hard boundary.

It also expands a direct internal 3-D A1 reference such as `Jan:Mar!B2` across every worksheet tab between its endpoints. Profiles identify those formulas. Since inserting, moving, or removing a tab can change that span without changing the formula text, a diff emits `FF014`; `no_3d_reference_scope_changes` can make that a hard boundary. This follows [Excel's documented 3-D-reference behavior](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets). External, malformed, endpoint-missing, and non-A1 variants remain explicit coverage limits rather than guessed dependencies.

Version 0.7.0 follows the static paths behind formula-defined names as well. A name such as `DiscountedValue` can itself use a second name such as `TaxRate`; when the complete definition resolves to internal dependencies or a constant, FormulaFence traces through to the source cells. It handles nested workbook and sheet-local definitions without evaluating the formulas. Relative, cyclic, external, dynamic, 3-D-in-name, and tokenizer-unsupported definitions remain explicit coverage gaps instead of guessed paths. That matches [Excel's support for names representing formulas and constants](https://support.microsoft.com/en-us/excel/names-in-formulas).

Version 0.8.0 removes a common modern false gap: `LET` names and inline `LAMBDA` parameters are local variables, not workbook names. FormulaFence follows their lexical scope—including nested lambdas supplied to functions such as `REDUCE`—and retains the actual A1, named, and table inputs around them. It follows Excel's documented [LET scope](https://support.microsoft.com/en-us/excel/functions/let-function) and [LAMBDA parameter syntax](https://support.microsoft.com/en-us/excel/functions/lambda-function), without evaluating either.

Version 0.9.0 follows a named `LAMBDA` only when its whole definition is static and internal. A caller such as `=ToCelsius(A2)` retains its `A2` edge and gains the function body's static inputs; nested named functions and formula-defined names that call one work the same way. It respects workbook and worksheet-local scope, and recognizes both ordinary Name Manager text and the `_xlfn.LAMBDA`/`_xlpm.`/`_xlop.` OOXML spelling emitted by Excel-compatible writers. That is the same named-function model described in [Excel's LAMBDA documentation](https://support.microsoft.com/en-us/excel/functions/lambda-function). Recursive, dynamic, relative, external, 3-D, and tokenizer-unsupported definitions remain explicit coverage notes rather than guessed graph edges.

Version 0.10.0 traces the anchor behind a direct internal spill reference such as `=SUM(A1#)` or its stored OOXML form `=SUM(_xlfn.ANCHORARRAY(A1))`. The anchor edge captures changes to its formula and visible upstream inputs; profiles retain the spill token, and `FF015`/`no_new_spill_references` let CI make the partial boundary explicit. It does not invent a dynamic spill extent or dependency on every possible blocker, so external, 3-D, range, named, implicit-intersection, and malformed spill forms stay visible limits. A spill-bearing formula-defined name remains unresolved at its caller. This follows Excel's [spilled-range operator](https://support.microsoft.com/en-us/excel/guidelines-and-examples-of-array-formulas), while recognizing the OOXML spelling used by [Excel-compatible writers](https://xlsxwriter.readthedocs.io/working_with_formulas.html). A formula the underlying tokenizer cannot inspect is now listed by location and emits `FF016`; `no_new_tokenization_failures` can make that a CI boundary too.

Version 0.11.0 makes the adjacent implicit-intersection boundary visible. Excel’s `@` operator can reduce a range or array to one value, and Excel-compatible files persist the unusual mixed cases as `_xlfn.SINGLE(...)`. When that wrapper contains one direct static A1 cell or range with an unambiguous row/column intersection, FormulaFence adds only the selected cell edge: `=_xlfn.SINGLE(Inputs!B2:B4)` in row 3 depends on `Inputs!B3`, not every cell in the range. It records literal `@A1:A3`, `@` applied to function results, and stored `SINGLE()` forms; a new use emits `FF017`, and `no_new_implicit_intersections` can require review. Function results retain their visible static inputs without being evaluated, and context-dependent formula-defined names remain unresolved. This is distinct from table `[@Column]` syntax. The scope follows Microsoft’s [implicit-intersection guidance](https://support.microsoft.com/en-us/excel/implicit-intersection-operator), its [Formula versus Formula2 documentation](https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2), and [XlsxWriter’s OOXML guidance](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

Version 0.12.0 closes a quieter graph gap in old but still material models. A legacy CSE formula can declare a fixed result range such as `B1:B3`, while the formula itself is stored only at `B1`. A downstream `=B2*10` or `=SUM(B2:B3)` therefore looks disconnected unless the tool understands the result members. FormulaFence keeps that fixed range compact and links its anchor to known formulas that read its members, so input changes reach those consumers without creating one virtual node per output cell. It identifies dynamic arrays through the workbook’s metadata rather than mistaking a current spill range for a fixed CSE range. Unrecognized metadata remains a coverage warning rather than a guess. The release also emits `FF018` when array mode or a fixed CSE output range changes; `no_array_formula_semantics_changes` makes that a CI boundary. The behavior follows [Microsoft’s distinction between dynamic and legacy arrays](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior) and [XlsxWriter’s documented storage forms](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

Version 0.13.0 closes the corresponding modern gap without pretending a dynamic spill is fixed. A dynamic-array anchor’s OOXML range is a current observation, not a promise about the next recalculation: Excel can grow, shrink, or block the spill. When a static formula currently reads a non-anchor member such as `B2`, FormulaFence now adds a compact anchor-to-consumer edge and records both the observed range and consumer in the profile. That lets an input change reach `=B2*10` and `=SUM(B2:B3)` through the dynamic anchor, without materializing every spill cell. A new observed output-member relationship emits `FF019`; `no_new_dynamic_array_output_references` turns it into the fail-closed `FFP019` policy boundary. The implementation was checked against an independently maintained XlsxWriter 3.2.9 workbook and keeps a declared `B1:XFD1048576` range compact.

Version 0.14.0 makes data-entry controls reviewable too. A worksheet validation can restrict a whole column, source a list from another sheet, show an input prompt, or block an invalid entry; all of those can materially change a model’s operating boundary without altering a formula cell. FormulaFence now inventories compact target ranges and effective validation settings, normalizes equivalent OOXML defaults, criterion spelling, and target grouping so compatible writers do not create noise, and emits `FF020` for a real change. Profiles redact validation formulas and prompt/error text; the local change report preserves full before/after evidence. Teams can make the surface fail closed with `no_data_validation_changes` (`FFP020`). It deliberately does not evaluate a validation formula or predict whether Excel will accept an entry. The behavior was checked against an independently maintained XlsxWriter workbook and follows [Microsoft’s data-validation guidance](https://support.microsoft.com/en-US/Excel/get-started/apply-data-validation-to-cells).

Version 0.15.0 makes visual exception controls reviewable too. Conditional formatting determines the colors, badges, and warnings a reviewer sees; moving a rule in worksheet-wide precedence, toggling `Stop If True`, changing a threshold, or changing a data bar can alter that signal without changing an ordinary formula cell. FormulaFence reads those rules directly from OOXML, resolves differential styles instead of comparing unstable `dxfId` values, and retains opaque Excel-2010 extension fragments when a reader library cannot model them. It normalizes equivalent defaults, leading `=` spelling, priority-number gaps, and known extension-link GUIDs, then emits `FF021` for a real change. Profiles redact criteria and raw visual XML; the local report keeps full before/after evidence. Teams can make the boundary fail closed with `no_conditional_formatting_changes` (`FFP021`). It does not evaluate a rule or predict Excel’s final rendered display. The behavior was checked against Microsoft’s public [conditional-formatting examples](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel) and an independently maintained XlsxWriter Excel-2010 data-bar fixture.

Version 0.16.0 extends that review surface to operational protection. A workbook can keep its formulas intact while losing a structure lock, allowing a protected input to become editable, hiding a formula differently, or changing who may edit a protected range. FormulaFence now reads workbook, worksheet, dialog-sheet, and chart-sheet protection directly from OOXML; records protected-range target areas; and inventories sparse direct `locked`/`hidden` assignments without expanding whole rows or columns. It normalizes equivalent worksheet-action defaults and emits `FF022` for a real change; `no_protection_changes` makes that a fail-closed `FFP022` boundary. Password verifiers, hashes, salts, protected-range names, and security descriptors never enter a profile or report—only safe presence metadata and private change fingerprints do. This is deliberately not a claim that Excel protection is encryption, authentication, or a full access-control system. The behavior was checked against Excel-produced public protection examples and controlled protected-range and chart-sheet fixtures.

It now traces common row-scoped forms without turning a row calculation into a dependency on every row of a table. Inside a table data cell, `[@[Sales Amount]]` and `[Sales Amount]` bind to that row. Qualified forms such as `Sales[@Amount]` and `Sales[[#This Row],[Amount]:[Rate]]` bind to the named table's data row even from an adjacent cell. That follows [Excel's documented structured-reference semantics](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables); header, total, cross-sheet, ambiguous, and complex bracket-escape cases remain coverage notes instead of guessed dependency paths.

One practical safeguard is coverage visibility. When the workbook parser encounters an OOXML extension it cannot fully interpret, FormulaFence records a coverage note. A candidate that adds one can be rejected with `no_new_parser_warnings`. Profiles also list unresolved range tokens, dynamic reference functions, spill references, explicit implicit intersection, dynamic and unclassified array anchors, observed dynamic output-member relationships, and formulas the tokenizer could not inspect; a change can be rejected with `no_new_unresolved_references`, `no_new_dynamic_references`, `no_new_spill_references`, `no_new_dynamic_array_output_references`, `no_new_implicit_intersections`, `no_array_formula_semantics_changes`, `no_data_validation_changes`, `no_conditional_formatting_changes`, `no_protection_changes`, or `no_new_tokenization_failures`.

## Test beyond toy files

Unit fixtures are necessary, but an Office-file reader also needs to meet real workbooks. I validated FormulaFence against the public [Foresight Cap Table and Exit Waterfall Tool](https://github.com/foresighthq/cap-table-tool): 18 sheets, 6,623 non-empty cells, and 4,228 formula cells.

The inspection found an unsupported OOXML extension and recorded it as a structured coverage note instead of leaking a raw dependency warning into CI output. It also identified 36 cells using `INDIRECT`, making the model’s dynamic-reference surface explicit. On a local, non-distributed copy, replacing one exit-waterfall formula with a number traced 330 downstream formula cells; the starter policy rejected both the formula override and the impact limit.

I also checked a public structured-reference workbook: FormulaFence identified its one table and 15 table-reference formulas with no unresolved tokens. Changing one table data cell on a local copy traced 16 downstream formulas, including a table total and an output outside the table. For the new row-scoped logic, a controlled workbook changed one input and traced exactly its calculated-row cell, a neighboring qualified row-reference cell, and an external summary—without marking the other two table rows. The [validation record](https://github.com/SybilGambleyyu/formulafence/blob/main/docs/validation.md) describes the boundary and evidence.

Those results are a compatibility demonstration, not a claim that the source model is correct. The full limits and validation record are in the [FormulaFence repository](https://github.com/SybilGambleyyu/formulafence/blob/main/docs/validation.md).

## What it does not claim

FormulaFence does not calculate Excel or prove a financial model correct. Material models still need a qualified owner, recalculation in the approved spreadsheet engine, and independent review. Relative, cyclic, external, and tokenizer-unsupported name definitions, unsupported or ambiguous table syntax, spill extents and blocking cells, non-static named LAMBDAs, arbitrary custom functions, complex implicit-intersection expressions, and features such as `INDIRECT` remain coverage limits rather than guessed graph edges.

But a review process should at least make it hard to silently replace a formula with a number. That is the narrow, useful boundary FormulaFence is built to enforce.

The current release is [FormulaFence 0.16.0 on GitHub](https://github.com/SybilGambleyyu/formulafence/releases/tag/v0.16.0). The canonical version of this post lives at [sybilgambleyyu.github.io/posts/formulafence.html](https://sybilgambleyyu.github.io/posts/formulafence.html).
