# LEVEL_3 exploration evaluation

All three Codex runs reported the correct **58,726,603** net receipts at `Summary!H18`. Their explanatory coverage differed: Sol recovered all required concepts, Terra omitted the row-count versus unit-quantity distinction, and Luna left several pricing, adjustment, and aggregation details unexplained despite obtaining the correct total.

| Model | Required-fact score | MCP calls | Failed calls | Exact repeat requests | Elapsed seconds |
| --- | ---: | ---: | ---: | ---: | ---: |
| `gpt-6.1-sol` | 28/28 | 68 | 4 | 0 | 101.962 |
| `gpt-6.1-terra` | 27/28 | 62 | 5 | 0 | 126.561 |
| `gpt-6-luna` | 21/28 | 61 | 14 | 2 | 81.946 |

Each model has one measured run. Scores assess the saved explanations against the specification, while tool failures and verification depth are reported separately. A correct headline total alone does not establish complete recovery of the workbook's rules.

## Conditions and saved evidence

The protocol matches LEVEL_1/LEVEL_2: fresh ephemeral sessions, **medium** reasoning effort, a **100-call** MCP ceiling, a **600-second** supervisor limit, and no additional token cap beyond model context/default output limits. Runs were sequential in Sol/Terra/Luna order. None received a supervisor hint, follow-up prompt, or a repeat exploration to improve its answer.

The exact task prompt, followed by the absolute LEVEL_3 workbook path, was:

> Explore this workbook using the supplied Excel MCP tools. Explain its purpose, sheet roles, inputs, reporting period if any, and how its final monetary total is calculated. Trace representative formulas to their sources and cite sheet/cell references. Describe refund, cancellation, pricing, and exception rules where present. Distinguish verified facts from assumptions and identify anything you could not establish.

The developer experiment constraints were unchanged. Each exploring process had an empty temporary working directory and access to the frozen workbook only through the five supplied MCP tools. Shell execution, other app/plugin sources, browsers, computer/image access, memory and delegation were disabled. User config/rules, project documents and automatic skills were disabled. The guarded stdio proxy forwarded requests to unchanged `mcp.py` and restricted them to this workbook. Code Mode's host remained enabled for MCP routing.

The specification, reference, canonical records, validation results, rubric, other models' answers and this report were available only to the supervising evaluator. They were not passed into the exploring sessions.

| Field | Sol | Terra | Luna |
| --- | --- | --- | --- |
| Exact model requested | `gpt-6.1-sol` | `gpt-6.1-terra` | `gpt-6-luna` |
| Start, JST on 2026-11-11 | 12:09:18.219 | 12:11:53.941 | 12:14:33.362 |
| End, JST on 2026-11-11 | 12:11:00.391 | 12:14:00.624 | 12:15:55.413 |
| Session | `01a1c894-d148-7b61-8d84-637452146457` | `01a1c897-3126-70f1-8e86-6708ea927d62` | `01a1c899-9fe4-7d83-b121-39a1939f75fe` |
| Termination | Normal, exit 0 | Normal, exit 0 | Normal, exit 0 |

All used Codex CLI **0.167.1**, Python **3.14.10**, MCP SDK **1.30.0**, and openpyxl **3.1.5**. Elapsed time uses the supervisor's monotonic clock and includes startup and export. Model identifiers and reasoning effort are exact client selections; an immutable backend model snapshot is not exposed by the event stream.

| Artifact | Sol | Terra | Luna |
| --- | --- | --- | --- |
| Unedited final explanation | [Markdown](exploration_codex_gpt_6_1_sol01.md) | [Markdown](exploration_codex_gpt_6_1_terra01.md) | [Markdown](exploration_codex_gpt_6_luna01.md) |
| Raw CLI event stream | [JSONL](exploration_codex_gpt_6_1_sol01.jsonl) | [JSONL](exploration_codex_gpt_6_1_terra01.jsonl) | [JSONL](exploration_codex_gpt_6_luna01.jsonl) |
| Timestamped MCP requests and results | [Audit](exploration_codex_gpt_6_1_sol01.mcp.jsonl) | [Audit](exploration_codex_gpt_6_1_terra01.mcp.jsonl) | [Audit](exploration_codex_gpt_6_luna01.mcp.jsonl) |
| Exact prompt, arguments, versions, budgets, hashes | [Metadata](exploration_codex_gpt_6_1_sol01.metadata.json) | [Metadata](exploration_codex_gpt_6_1_terra01.metadata.json) | [Metadata](exploration_codex_gpt_6_luna01.metadata.json) |
| Derived metrics and evidence checks | [Metrics](exploration_codex_gpt_6_1_sol01.metrics.json) | [Metrics](exploration_codex_gpt_6_1_terra01.metrics.json) | [Metrics](exploration_codex_gpt_6_luna01.metrics.json) |
| CLI diagnostics | [Log](exploration_codex_gpt_6_1_sol01.stderr.log) | [Log](exploration_codex_gpt_6_1_terra01.stderr.log) | [Log](exploration_codex_gpt_6_luna01.stderr.log) |

The [runner](../experiments/run_codex_exploration.py) was extended only to accept `--level 3`; prompt, budgets and restrictions are unchanged. The exact prior LEVEL_2 runner is preserved in [run_codex_exploration_level2.py](../experiments/run_codex_exploration_level2.py), matching those earlier metadata hashes. Existing LEVEL_1/LEVEL_2 results were retained unchanged.

## Independent validation and frozen input

The evaluator ran `.venv/bin/python LEVEL_3/validation.py --require-caches`, saving its JSON output in [exploration_validation.json](exploration_validation.json). It passed source/structure checks and compared **440,370 saved formula results** with independent policy calculations. There were **zero missing caches**, and `experiment_ready` was true. This exhaustive evaluator work was not performed by the models or counted as their MCP usage.

All sessions recorded the same workbook hash before and after:

```text
a4a9c154cf4cb67391a96450a2401be43734e8de51ca07fb02c4069b3cff240f
```

It matches the recalculated LEVEL_3 entry in [recalculation_manifest.json](../recalculation_manifest.json), which records Microsoft Excel **16.114.4**. The original [reference.json](reference.json) retains the pre-recalculation generation hash and status; the later manifest and strict validation establish the frozen input's readiness. `mcp.py` and `tools.py` were unchanged throughout.

The [offline checker](../experiments/evaluate_level3_trace.py) reconstructs cells from the recorded CSV/formula results and compares captured values with the evaluator's integer-arithmetic oracle using canonical records. It confirms that oracle totals agree with the reference. Formula text is compared with prescribed templates, normalizing Excel's `_xlfn.`/`_xlpm.` prefixes. It checks only sampled Calc/Transactions observations and reports unobserved evidence separately; it does not open the workbook or make MCP/model calls. For example:

```bash
.venv/bin/python experiments/evaluate_level3_trace.py --run-id exploration_codex_gpt_6_1_sol01
```

## Required-fact scoring

The [rubric](../experiments/level3_rubric.json) was saved before observing LEVEL_3 explanations. It follows [the specification and protocol](../excel_book_levels_specification.md): six LEVEL_2 concept groups, adapted to LEVEL_3, plus eight added concepts. Each receives **0** for missing/incorrect, **1** for partial/unsupported, or **2** for correct with relevant workbook evidence. The maximum is **28**. Numeric scores across levels have different denominators.

| Required fact | Sol | Terra | Luna | Evidence and assessment |
| --- | ---: | ---: | ---: | --- |
| Eleven sheet roles | 2 | 2 | 2 | All identify raw Transactions, normalized Calc, the five added support sheets, status reports and Summary. |
| Reporting boundaries | 2 | 2 | 2 | All explain the January start-inclusive/end-exclusive window and cite `Summary!B2:C2`. |
| Final-state row semantics | 2 | 2 | 1 | Sol/Terra explicitly describe synthetic final-state order lines. Luna explains conditional branches on a transaction row but leaves the one-line/final-state interpretation implicit. |
| Status quantities versus row counts | 2 | 1 | 0 | Sol explains aggregated quantities and matching-row counts. Terra describes product/status sums but omits the count/unit distinction. Luna does not explain the status-specific quantity sums or transaction counts in its final response. |
| Charges, refunds and cancellation treatment | 2 | 2 | 2 | All include original charges on sold/refunded lines, refund subtraction and fees, with canceled initial charges zero. |
| Net receipts and final value | 2 | 2 | 2 | All explain `InitialCharges-RefundsPaid+CancellationFees+Adjustments` and report 58,726,603 with Summary references. |
| Alias normalization and missing mappings | 2 | 2 | 1 | All explain product/status maps. Sol/Terra also identify `#N/A` for missing mappings. Luna omits exact/error-fallback behavior. |
| Effective-dated prices | 2 | 2 | 1 | Sol explains V1/V2, the January 16 transition, 10% increase and unchanged fees. Terra gives the date predicate and P001's representative 1,000→1,100 transition with unchanged fees. Luna explains date-based selection but omits the transition date and increase parameters. All distinguish PriceHistory from Master reference prices. |
| Channel discounts | 2 | 2 | 2 | All explain DIRECT 0%, PARTNER 5%, VIP 10% and goods-only discounting through the row formulas. |
| Shipping-refund branches | 2 | 2 | 2 | All identify VIP's shipping refund and DIRECT/PARTNER's retained shipping; refunds return discounted goods. |
| Cancellation multipliers | 2 | 2 | 2 | All give DIRECT 1, PARTNER 1.5, VIP 0 and apply the multiplier to quantity × cancellation fee. |
| Per-row rounding | 2 | 2 | 2 | All locate goods and fee rounding in per-transaction Calc formulas before the reports consume the calculated amounts. Sol additionally demonstrates an actual fractional goods amount. |
| Signed adjustments | 2 | 2 | 1 | Sol/Terra explain ±500 ID-based exceptions, zero for absent IDs and inclusion in the final calculation. Luna identifies the lookup and named ranges but omits the entry amounts and missing-ID fallback; it never successfully reads the adjustment values. |
| Complete aggregate dependency path | 2 | 2 | 1 | Sol/Terra explicitly connect status aggregates to Calc row amounts and their source/price/rule dependencies. Luna explains Calc and Summary but leaves the intervening status-sheet monetary `SUMIFS` path and its significance unclear. |
| **Total** | **28/28** | **27/28** | **21/28** | **Missing explanation is distinguished from a wrong numeric claim.** |

Full credit recognizes the mechanism explained and supporting representative evidence; it does not require transcribing every formula or independently auditing every source row. Terra's representative price example is sufficient to demonstrate the historical-price mechanism. Its row-count omission is distinct because no explanation identifies counts as rows rather than units. Likewise, Luna receives partial credit for discovered but incompletely explained policies, rather than being treated as having missed those sheets entirely.

Representative formula evidence is in Sol's `item_27`–`item_42` and `item_59`–`item_62`; Terra's `item_48`–`item_64`; and Luna's successful `item_28`, `item_31`–`item_33`, `item_47`–`item_52` and `item_60`. Each trace's `get_named_regions` result exposes all 17 lookup names. The final-answer scores are evaluator judgments, separate from the mechanical trace checker.

## Numeric accuracy and verification depth

The independent grand-total bridge is:

| Component | Cell | Expected | Sol | Terra | Luna |
| --- | --- | ---: | ---: | ---: | ---: |
| Initial charges | `Summary!E18` | 74,059,268 | 74,059,268 | 74,059,268 | 74,059,268 |
| Refunds paid, subtracted | `Summary!F18` | 16,013,665 | 16,013,665 | 16,013,665 | 16,013,665 |
| Cancellation fees | `Summary!G18` | 681,000 | 681,000 | 681,000 | 681,000 |
| Signed adjustments, net | `Summary!K18` | 0 | 0 | 0 | 0 |
| **Net receipts** | **`Summary!H18`** | **58,726,603** | **58,726,603** | **58,726,603** | **58,726,603** |

All three also correctly describe the zero reconciliation total. Sol and Terra report the correct P001 bridge: **1,914,435 − 410,120 + 14,450 + 500 = 1,519,265**. No incorrect reported monetary amount was found. Zero net adjustments are not evidence of absent adjustment rules: product-level adjustments are nonzero and the positive/negative entries offset overall.

The offline checker found **zero mismatches** in captured observations. It matched **1,693 / 1,143 / 885** checked values for Sol/Terra/Luna and **181 / 236 / 276** captured formula cells respectively. These counts mix small support/report tables with sampled Calc/source observations; they are not counts of fully verified transactions. Unobserved small-table cells and repeated formulas remain listed in the metrics rather than being treated as failures.

Sol gives several concrete examples beyond describing formula patterns:

- Row 2 normalizes `SKU007`/`PAID`, selects the February P007 V2 price and produces 23,500, explicitly outside the January report.
- Row 10 rounds a PARTNER P006 sale from `5*3850*0.95 = 18,287.5` to **18,288** goods, then adds 1,000 shipping to reach **19,288**.
- Row 12 demonstrates a DIRECT January refund retaining **800** shipping.
- Row 20 demonstrates a DIRECT cancellation yielding **1,350**.
- Row 25 demonstrates a VIP refund returning both goods and shipping, explicitly outside January.

Terra provides a historical-price example and P001 report bridge but does not independently demonstrate a rounded transaction or a nonzero adjusted source row. Luna gives a correct February row-2 example but no concrete transition, rounding-boundary, or nonzero adjustment example.

None reads the specific guaranteed edge-case rows listed in the evaluator reference, including the exact price transition at row 3133, the adjustment at row 15239, and the VIP refund at row 19440. Sol's alternative sampled examples still support several policy branches. Correct formula explanations and matching samples do not prove global mapping uniqueness, complete historical interval coverage, or that every transaction has correct caches. Those stronger checks were performed by the separate validator.

All three warn that MCP results are saved caches rather than live recalculation. Sol explicitly distinguishes formula/sample inspection from a full audit and notes unresolved duplicates/overlaps/invalid inputs. Terra notes compatibility, cash-timing and formula-range limits. Luna gives cache/Excel-function caveats but less detail about unverified data integrity. None establishes currency or timezone from the exposed cell values. Categorical claims that the workbook does not identify currency should be narrowed to currency not being identified in those values, since number formats are not exposed.

## Observable strategies and tool errors

**Sol** discovers all names/sheets and bounds all 11 sheets with `A1:AZ100000` searches. It first reads the top-left windows, then follows Summary and Calc formulas into high columns and named sources. It reads all 20 adjustment records as values, all product aliases, all price-history rows and all channel policies. Its source sample covers rows **2–30 and 19992–20001** (39 rows). Calc values cover 32 rows, with only rows **2–10** read completely across all 22 columns. Its formula samples include the final three Calc rows. No exact request is repeated.

Sol's four failed value requests are `Summary!A11:K18` (11 columns), `Rules!K11:N21` (11 rows), `Summary!K5:K18` (14 rows), and `Calc!U11:V30` (20 rows). It repairs each by splitting the ranges. These failures count toward the 68 calls; no supervisor assistance was supplied.

**Terra** initially searches `A1:AZ10000`. It later extends the Transactions and Calc searches to `A10001:AZ100000`, resolving their actual row-20001 extent. It reads the complete small support/report values, including all 20 adjustment records. Its source values cover only rows **2–10**, and Calc values only **2–5**, but it successfully reads all 22 formulas at row 20001. Its five oversized requests are `Summary!A11:K18`, `Rules!A11:N20`, `Transactions!A20000:H20010`, `Calc!A20000:J20010`, and the formula range `Summary!K1:K18`. It recovers the policy/report evidence and late-row formula evidence without exact repeats.

**Luna** searches only `A1:AZ200`, reads source rows **2–10 and 191–200**, and never establishes the full 20,000-row extent through a wider search or late-row read. Its Calc value observations span nine rows, but it never reads a complete 22-column Calc row as values. It obtains the header/mapping formulas and a few monetary columns separately. It reads the full Summary values eventually, while portions of the master, alias and price-history tables remain unobserved.

Luna's 14 rejected ranges are:

| Trace item | Tool | Range |
| --- | --- | --- |
| `item_26` | Values | `Calc!A1:V10` |
| `item_29` | Formulas | `Summary!A1:K18` |
| `item_30` | Formulas | `Calc!K1:V10` |
| `item_35` | Values | `Calc!K199:V200` |
| `item_36` | Values | `Calc!K1:V3` |
| `item_39` | Values | `Rules!A11:N21` |
| `item_40` | Values | `Summary!A1:K18` |
| `item_43` | Formulas | `Summary!A1:K10` |
| `item_44` | Formulas | `Calc!K2:V2` |
| `item_45` | Formulas | `Calc!K3:V3` |
| `item_54` | Values | `Summary!A11:K18` |
| `item_57` | Values | `Rules!K11:N21` |
| `item_58` | Values | `Calc!A1:V1` |
| `item_59` | Values | `Calc!K4:V4` |

All fail the documented 10-row/10-column constraint. Some contain just one row but still exceed the column limit. Luna eventually splits the important Calc monetary ranges into `K:T` and `U:V`, and Summary into `A:J` and `K`. It does not successfully repair the adjustment value reads. Instead, it makes valid **formula** requests on `Rules!K1:N10`, `K2:N5` and `K11:N20`; these return `{}` because the adjustment block contains constants. It does not follow with valid value requests there. Its two exact repeats are successful Summary value reads (`A1:J10` and `K11:K18`).

Luna therefore discovers the adjustment branch's existence through defined names and Summary formulas but never inspects its individual IDs, signed amounts, reasons or approval references. Its final statement that the workbook does not explain why entries were issued is not supported by having read the Reason fields; the block actually contains synthetic reason and approval-reference labels. Those labels do not independently establish a real business justification, either.

## Cost and interpretation

| Measure | Sol | Terra | Luna |
| --- | ---: | ---: | ---: |
| Total supplied MCP calls | 68 | 62 | 61 |
| Successful / failed | 64 / 4 | 57 / 5 | 47 / 14 |
| Sheet / name / area calls | 1 / 1 / 11 | 1 / 1 / 13 | 1 / 1 / 11 |
| Value / formula calls, including failures | 38 / 17 | 32 / 15 | 31 / 17 |
| Exact repeated requests | 0 | 0 | 2 |
| Proxy rejections | 0 | 0 | 0 |
| Defined names returned | 17 | 17 | 17 |
| Distinct raw source rows with values | 39 | 9 | 19 |
| Distinct Calc rows with any values | 32 | 4 | 9 |
| Fully read 22-column Calc value rows | 9 | 4 | 0 |
| Adjustment records read as values | 20 | 20 | 0 |
| Value cells requested successfully, including blanks/repeats | 2,100 | 1,243 | 1,406 |
| Unique nonempty value cells returned | 1,703 | 1,153 | 895 |
| Unique formula cells returned | 181 | 236 | 276 |
| CLI input tokens | 143,678 | 117,518 | 204,286 |
| Cached input tokens, included above | 107,136 | 89,344 | 173,440 |
| Cache-write input tokens | 0 | 0 | 0 |
| CLI output tokens | 3,621 | 2,496 | 3,229 |
| Separately reported reasoning-output counter | 347 | 72 | 187 |
| Elapsed seconds | 101.962 | 126.561 | 81.946 |

Counters are reported as emitted by the CLI, not converted into a cash bill or used to infer private reasoning. Input tokens accumulate context across requests. Formula/value reads of the same rectangle serve different purposes and are not exact repeats. Rejected ranges return no cell data and do not count as successful returned-value areas.

There were no setup failures, supervisor interventions or budget exhaustion in these LEVEL_3 runs. All emitted the same nonfatal warning about the under-development `skip_host_skill_discovery` feature. The shorter Luna run reflects less completed evidence collection as well as different execution conditions and cannot be interpreted as a general performance advantage.

For the blog hypothesis, the evidence supports a specific conclusion: Sol successfully recovered the required LEVEL_3 logic; Terra recovered nearly all of it; Luna encountered repeated tool-constraint problems and left several policy details and the adjustment data unresolved. All three still reported the correct total. These observations do not support saying that all models struggled with LEVEL_3, nor do they establish reliability without repeated trials.
