# Excel exploration evaluation

All three LEVEL_1 explorations recovered the five required calculation facts, scoring **10/10** each. Luna initially exceeded the read-size limit on four calls, then recovered and completed correctly. Factual accuracy and tool-use errors are reported separately.

| Model | Required-fact score | Elapsed seconds | MCP calls, including failures | Failed MCP calls |
| --- | ---: | ---: | ---: | ---: |
| `gpt-6.1-sol` | 10/10 | 54.418 | 12 | 0 |
| `gpt-6.1-terra` | 10/10 | 32.577 | 12 | 0 |
| `gpt-6-luna` | 10/10 | 37.973 | 16 | 4 |

Each model has one measured workbook exploration. These observations do not establish a general model ranking or run-to-run reliability. [The Terra/Luna comparison](#terra-and-luna-comparison) follows the original Sol evaluation below.

## Sol baseline

Codex with **`gpt-6.1-sol` successfully recovered the required LEVEL_1 calculation logic: 10/10 points across five criteria**. It correctly identified the inputs and exact-match master lookups, explained quantity-based shipping, and reported the correct subtotal and grand-total formulas and values. No incorrect monetary amounts or substantive calculation claims were found.

This baseline is one successful workbook exploration, not evidence about reliability across repeated runs or about other levels or models.

### Run and saved evidence

| Field | Recorded value |
| --- | --- |
| Run | `exploration_codex_gpt_6_1_sol01` |
| Agent | Codex CLI `0.167.1` |
| Model selection | Explicit CLI argument `--model gpt-6.1-sol`; no substitute model requested |
| Reasoning setting | Explicit `model_reasoning_effort="medium"` |
| Workbook | `LEVEL_1/excel_level_1.xlsx` |
| Start / end | 2026-11-11 11:27:43.607 / 11:28:38.118 JST |
| Elapsed time | 54.418 seconds, including CLI startup and final export |
| Termination | Normal completion, process exit code 0 |
| Session | `01a1c86e-c037-7f90-b2bf-7efd694664ef` |
| Predeclared budgets | 100 MCP calls; 600 seconds; no additional token cap beyond model context and default output limits |
| MCP environment | Python 3.14.10; MCP SDK 1.30.0; openpyxl 3.1.5 |

The model identifier and reasoning setting above are the exact client selections. The exported events do not independently identify an immutable backend model snapshot.

The primary artifacts are:

- [Exploration explanation](exploration_codex_gpt_6_1_sol01.md): Codex's final response, saved directly with `--output-last-message`, without evaluator edits.
- [Raw exploration trace](exploration_codex_gpt_6_1_sol01.jsonl): the unmodified `codex exec --json` event stream, including MCP arguments and results.
- [MCP audit](exploration_codex_gpt_6_1_sol01.mcp.jsonl): timestamped JSON-RPC requests and responses, including tool discovery.
- [Run metadata](exploration_codex_gpt_6_1_sol01.metadata.json): exact prompt, command arguments, settings, versions, budgets, hashes, and usage.
- [Derived metrics and numeric checks](exploration_codex_gpt_6_1_sol01.metrics.json).
- [CLI diagnostics](exploration_codex_gpt_6_1_sol01.stderr.log).

The original runner is preserved in [run_codex_exploration_sol01.py](../experiments/run_codex_exploration_sol01.py); its contents match the harness hash recorded for Sol. The original evaluator-only checks are preserved in [evaluate_level1_trace_sol01.py](../experiments/evaluate_level1_trace_sol01.py) and can be reproduced with `.venv/bin/python experiments/evaluate_level1_trace_sol01.py`. The active runner and checker have since been parameterized for Terra and Luna, as described below. Existing Sol logs, metadata, explanation, and metrics were retained unchanged.

### Experimental separation

The exploring Codex process started a fresh, ephemeral session in an empty temporary directory. It received the protocol prompt below and the workbook path, plus experiment constraints in a developer instruction. It received neither the specification nor the evaluator reference, generator source, blog plan, or this report. The supervising/evaluating agent read those files separately.

Shell execution, other app/plugin integrations, browser/computer access, image access, memory, and delegation were disabled. User configuration and execution rules were ignored for this run, and project-document loading and automatic skills were disabled. A proxy forwarded requests to the **unchanged supplied `mcp.py`**, allowing only its five Excel tools and only the assigned workbook path. The original tools enforced the 10 × 10 value/formula limit. The proxy also enforced the 100-call ceiling; the supervisor enforced the time limit.

Code Mode's host remained enabled because this CLI version requires it to route MCP calls. It served as the tool gateway. The recorded workbook access consists entirely of the 12 supplied MCP calls listed below; no alternative workbook access or evaluator-file access is recorded.

The exact user task was:

> Explore this workbook using the supplied Excel MCP tools. Explain its purpose, sheet roles, inputs, reporting period if any, and how its final monetary total is calculated. Trace representative formulas to their sources and cite sheet/cell references. Describe refund, cancellation, pricing, and exception rules where present. Distinguish verified facts from assumptions and identify anything you could not establish.

This was followed by `Workbook:` and the absolute LEVEL_1 workbook path. The complete developer constraint text and command are in the metadata.

The workbook SHA-256 before and after exploration was identical:

```text
1a507f80bad7dc5f7d6567a83a53b51b73ce5838e016992d78828036e76f8b3d
```

It matches the recalculated LEVEL_1 hash in [recalculation_manifest.json](../recalculation_manifest.json), which records Microsoft Excel 16.114.4. The evaluator reference still describes the earlier, unrecalculated generation hash and pending recalculation; the later manifest resolves that provenance difference. Both `mcp.py` and `tools.py` also retained their original hashes. No workbook was edited or recalculated during exploration.

### Evaluation against the specification

The rubric follows [the LEVEL_1 specification and experiment protocol](../excel_book_levels_specification.md). Scores are 0 = missing/incorrect, 1 = partially correct/unsupported, and 2 = correct with relevant workbook evidence. The protocol's required facts are expanded into five criteria below. Final-answer coverage is scored using the saved explanation, with the raw trace checking its evidence.

| Required fact | Expected specification | Evidence in the exported exploration | Score |
| --- | --- | --- | ---: |
| Inputs | Product codes in `Invoice!A5:A12` and quantities in `C5:C12`; eight purchase lines | Identifies both input ranges and eight products P001–P008. Separately identifies invoice ID/date and master constants. The trace reads all eight lines (`item_8`, `item_9`). | 2 |
| Master lookups | Exact-match lookup of product name, unit price, and shipping rate from `Master!A2:F13` | Explains the three lookup columns and `FALSE` exact matching; quotes `D5` and `E5` formulas and traces them to `Master!C2` and `D2`. All copied lookup formulas are present in `item_12` and `item_13`. | 2 |
| Quantity-based shipping | `G5=C5*E5`, copied through row 12; line totals combine goods and shipping | Explicitly describes per-unit charging and the P001 calculation `2*200=400`; cites a zero-shipping P008 example. Also explains `F5=C5*D5` and `H5=F5+G5`. | 2 |
| Subtotal formulas and values | `H14=SUM(F5:F12)` = 42,000; `H15=SUM(G5:G12)` = 2,300 | Reproduces both formulas and values with cell references, supported by the formula and cached-value reads. | 2 |
| Grand-total formula and value | `H16=SUM(H5:H12)` = 44,300; also equals `H14+H15` | Reproduces the actual formula and value. Correctly distinguishes summing line totals from directly referencing the two subtotal cells. The two subtotals numerically reconcile to the grand total. | 2 |
| **Total** | **Five required criteria** | **All five correctly explained with workbook evidence** | **10/10** |

Additional specification details were also recovered correctly: `Master` and `Invoice` roles; 12 master products; the invoice's ID/date cells; cancellation fees present in the master but unused in the invoice; and no refund calculations. It appropriately declined to infer a January reporting interval from a 31 January issue date.

The final prose does not enumerate the complete specified quantity vector `2,1,3,2,1,2,1,3`; those values are all present in the trace. It also does not reproduce every product-master row in its explanation. These are differences between a concise explanation and a full specification transcription, rather than missing required calculation concepts.

One specification convention remains undiscovered: **JPY**. Codex explicitly says no currency is specified in the inspected cell values. That is an appropriate limitation of these tools, which return underlying values and do not expose number formats. It should not be counted as an incorrect currency claim. The run likewise does not establish the specification's timezone convention or positive-integer validation as enforced workbook restrictions.

### Numeric accuracy and verification depth

The evaluator independently computed these amounts from the specification's quantities, unit prices, and per-unit shipping rates, then compared them with the CSV results preserved in the trace and the independent [evaluator reference](evaluator_reference.json).

| Result | Expected | Agent reported | Difference |
| --- | ---: | ---: | ---: |
| First line, `Invoice!H5` | 2,400 | 2,400 | 0 |
| Last line, `Invoice!H12` | 13,500 | 13,500 | 0 |
| Goods subtotal, `Invoice!H14` | 42,000 | 42,000 | 0 |
| Shipping total, `Invoice!H15` | 2,300 | 2,300 | 0 |
| Grand total, `Invoice!H16` | 44,300 | 44,300 | 0 |

The evaluator's automated checks passed for all 12 master rows, unique master product keys, all eight invoice rows and their calculated amounts, all 51 invoice formulas, and all three saved totals. These checks operate on the frozen exploration trace, without reopening the workbook or making further model calls.

Codex's final response states that arithmetic agrees with the saved totals. The visible evidence establishes that it obtained every required input and formula and produced correct examples and totals. It does not expose a separate calculation execution proving every intermediate arithmetic operation performed by the model. The exhaustive arithmetic comparison above is an **evaluator check**, distinct from the agent's statement.

The explanation correctly warns that MCP values are saved caches, with no live recalculation. It also distinguishes inferred editable inputs from verified edit permissions or validation. Its prediction that an unmatched product code would cause a propagating lookup error is supported by the formulas, but was not tested by modifying inputs.

### Observable exploration behavior and tool cost

The request order was:

| Order | Trace item | Tool | Sheet / range | Observation |
| ---: | --- | --- | --- | --- |
| 1 | `item_2` | `get_sheet_names` | Workbook | Finds `Master` and `Invoice`, in that order. |
| 2 | `item_3` | `get_named_regions` | Workbook | Finds no defined names. |
| 3 | `item_4` | `get_value_area` | `Invoice!A1:ZZ10000` | Returns `A2:H16`. |
| 4 | `item_5` | `get_value_area` | `Master!A1:ZZ10000` | Returns `A1:F13`. |
| 5 | `item_6` | `get_cell_values_csv` | `Master!A1:F10` | Reads headers and the first master records. |
| 6 | `item_7` | `get_cell_values_csv` | `Master!A11:F13` | Completes the master. |
| 7 | `item_8` | `get_cell_values_csv` | `Invoice!A1:H10` | Reads invoice metadata, headers, and six lines. |
| 8 | `item_9` | `get_cell_values_csv` | `Invoice!A11:H16` | Reads the last two lines and totals. |
| 9 | `item_10` | `get_cell_formulas` | `Master!A11:F13` | Confirms no formulas in this block. |
| 10 | `item_11` | `get_cell_formulas` | `Master!A1:F10` | Confirms the rest of the master is constant. |
| 11 | `item_12` | `get_cell_formulas` | `Invoice!A1:H10` | Obtains the first six lines' formulas. |
| 12 | `item_13` | `get_cell_formulas` | `Invoice!A11:H16` | Obtains the remaining formulas and totals. |

This is structure discovery followed by source-first value reading, then formula inspection across both sheets. It did not begin by tracing only the displayed grand total. Its small windows collectively cover all populated cell data in the two discovered rectangles, rather than sampling a few records. Multiple requests were issued before preceding responses completed; the table records request order, not a claim of strictly sequential execution.

The final explanation reaches the full dependency path: `Invoice!H16` through line totals, goods/shipping amounts, quantities and lookup outputs, to the master rates. Its P008 example checks the zero-shipping branch as well as the ordinary P001 case.

| Cost measure | Successful run |
| --- | ---: |
| Total supplied MCP tool calls | 12 |
| Sheet / name / area calls | 1 / 1 / 2 |
| Value / formula calls | 4 / 4 |
| Exact repeated tool-and-argument requests | 0 |
| MCP errors / proxy rejections | 0 / 0 |
| Requested value cells, including blanks | 206 |
| Nonempty value cells returned | 160 |
| Formula cells returned | 51 |
| CLI-reported input tokens | 68,556 |
| Cached input tokens, included in input count | 44,288 |
| Cache-write input tokens | 0 |
| CLI-reported output tokens | 1,712 |
| Separately reported reasoning-output counter | 115 |

Token counters are reported as emitted by the CLI, not inferred from workbook size or converted into a monetary bill. Input totals include model context across requests. The four formula reads revisit the same rectangles as the value reads for different information; they are not duplicate requests. The broad `A1:ZZ10000` area searches are allowed by the protocol, although bounds outside those search rectangles were not examined.

### Setup failures and interpretation limits

Two setup attempts were retained separately, before the successful run:

1. [Setup failure 01](../experiments/setup_failure_01/exploration_codex_gpt_6_1_sol01.stderr.log): the sandbox prevented local app-server initialization. No session started and no workbook reads occurred. The subsequent CLI launch used approved execution outside that sandbox.
2. [Setup failure 02](../experiments/setup_failure_02/exploration_codex_gpt_6_1_sol01.jsonl): disabling the Code Mode host prevented tool routing. Codex accurately reported that it could not inspect the workbook. The MCP audit contains initialization/discovery only, with **zero workbook tool calls**. This attempt took 20.747 seconds and reported 27,467 input tokens, 20,352 cached input tokens, 354 output tokens, and a reasoning-output counter of 153. Its metadata and response remain in the same setup-failure directory.

After restoring the required gateway host, the successful run used a new session and unchanged prompt/budgets. These setup failures are excluded from the workbook-discovery score and the successful-run cost table, and should remain visible in any account of operational setup. The successful session also emitted a nonfatal warning that `skip_host_skill_discovery` is an under-development feature; all 12 workbook calls succeeded.

The evidence supports describing this LEVEL_1 run as successful. It contains no unresolved calculation dependency, incorrect total, repeated failed workbook request, or budget exhaustion. It does not establish behavior under edits, live Excel recalculation, validation/protection settings, or unexamined business policies. Observable tool actions and the exported explanation support the findings; no private reasoning is inferred. Use the same prompt and predeclared budget policy for subsequent comparisons, and repeat workbook explorations to assess variability.

For the export mechanism, official OpenAI documentation describes [JSONL events and final-message export](https://learn.chatgpt.com/docs/non-interactive-mode). The selected identifier is documented on the [GPT-6.1 Sol model page](https://developers.openai.com/api/docs/models/gpt-6.1-sol). These documentation checks were performed by the supervising agent and were not supplied to the exploring model.

## Terra and Luna comparison

### Conditions and artifacts

Terra and Luna ran sequentially in separate fresh, ephemeral sessions. Both used the same user prompt, developer experiment constraints, medium reasoning setting, 100-call ceiling, 600-second time limit, and token-budget policy as Sol. The CLI and MCP dependency versions were identical. The input hashes before and after each run match Sol's frozen workbook, `mcp.py`, and `tools.py` hashes. Each process received only the assigned workbook path and accessed its contents through the five supplied MCP tools. The supervising agent's specification, references, model-documentation checks, and previous results were not passed to either session.

| Field | Terra | Luna |
| --- | --- | --- |
| Exact client model selection | `gpt-6.1-terra` | `gpt-6-luna` |
| Reasoning effort requested | `medium` | `medium` |
| Agent | Codex CLI 0.167.1 | Codex CLI 0.167.1 |
| Start, JST on 2026-11-11 | 11:49:46.694 | 11:50:38.117 |
| End, JST on 2026-11-11 | 11:50:19.747 | 11:51:16.185 |
| Session | `01a1c882-f31a-7a81-8ad5-45fa3d1f9e11` | `01a1c883-b970-7482-83b6-9897651f468e` |
| Termination | Normal completion; exit 0 | Normal completion; exit 0 |

Elapsed durations use the supervisor's monotonic clock. The wall-clock timestamps are retained separately. These exact model identifiers are documented by OpenAI for [GPT-6.1 Terra](https://developers.openai.com/api/docs/models/gpt-6.1-terra) and [GPT-6 Luna](https://developers.openai.com/api/docs/models/gpt-6-luna); an immutable served backend snapshot is not exposed in the run events.

| Artifact | Terra | Luna |
| --- | --- | --- |
| Unedited final explanation | [Markdown](exploration_codex_gpt_6_1_terra01.md) | [Markdown](exploration_codex_gpt_6_luna01.md) |
| Raw CLI event stream | [JSONL](exploration_codex_gpt_6_1_terra01.jsonl) | [JSONL](exploration_codex_gpt_6_luna01.jsonl) |
| Timestamped MCP audit | [MCP JSONL](exploration_codex_gpt_6_1_terra01.mcp.jsonl) | [MCP JSONL](exploration_codex_gpt_6_luna01.mcp.jsonl) |
| Prompt, settings, versions, hashes | [Metadata](exploration_codex_gpt_6_1_terra01.metadata.json) | [Metadata](exploration_codex_gpt_6_luna01.metadata.json) |
| Derived trace and numeric checks | [Metrics](exploration_codex_gpt_6_1_terra01.metrics.json) | [Metrics](exploration_codex_gpt_6_luna01.metrics.json) |
| CLI diagnostics | [Log](exploration_codex_gpt_6_1_terra01.stderr.log) | [Log](exploration_codex_gpt_6_luna01.stderr.log) |

The active [runner](../experiments/run_codex_exploration.py) now accepts `--model gpt-6.1-terra` or `--model gpt-6-luna` and gives each run and MCP audit its own output stem. It continues to refuse overwriting existing runs. These are model/path parameter changes; the experiment prompt, access restrictions, and budgets are unchanged.

The active [offline checker](../experiments/evaluate_level1_trace.py) accepts `--run-id exploration_codex_gpt_6_1_terra01` or `--run-id exploration_codex_gpt_6_luna01`. It now reports missing observations separately from incorrect observations and handles failed tool calls. It does not require exhaustive reads to award a narrative score, and it makes no new MCP or model calls. For these two small-workbook runs, all 139 checked master/invoice values and all 51 expected formulas happened to be observed and matched. The original Sol artifacts and scripts are preserved separately.

### Required-fact scoring

The same five criteria and 0–2 scoring scale used for Sol apply. Tool-use failures do not reduce a correct factual explanation's score; they appear in the cost and behavior assessment.

| Criterion | Terra explanation and supporting trace | Score | Luna explanation and supporting trace | Score |
| --- | --- | ---: | --- | ---: |
| Inputs | Identifies eight product lines, `A5:A12` codes and `C5:C12` quantities, invoice ID/date, and the master. Complete value reads are in `item_7`, `item_8`, `item_10`, `item_12`. | 2 | Identifies product codes and quantities as inputs, distinguishes the formula-driven columns, and cites the first line `A5:C5`. Complete value reads are in `item_12`, `item_13`, `item_15`, `item_17`. | 2 |
| Master lookups | Explains exact matching, quotes `D5`'s `VLOOKUP(...,3,FALSE)`, and traces price/shipping to `Master!C2:D2`. Formula reads `item_11`, `item_13` support all three lookup columns. | 2 | Explains exact-match `VLOOKUP` with `FALSE`, identifies master columns 3 and 4 for price/shipping, and distinguishes the unused cancellation column. Formula reads `item_16`, `item_18` support the lookup paths. | 2 |
| Quantity-based shipping | Gives `G5=C5*E5`, `2*200=400`, goods and line-total formulas, plus P004/P008 zero-shipping examples. | 2 | Gives `G5=C5*E5`, the 400 shipping result, and zero-shipping references `Master!D5`, `D9`; also explains goods and line totals. | 2 |
| Subtotals | Correctly states `H14=SUM(F5:F12)` = 42,000 and `H15=SUM(G5:G12)` = 2,300, supported by the lower invoice value/formula blocks. | 2 | Correctly states the same formulas and cached values, supported by `item_17`, `item_18`. | 2 |
| Grand total | Correctly states `H16=SUM(H5:H12)` = 44,300, supported by `item_12`, `item_13`. | 2 | Correctly states `H16=SUM(H5:H12)` = 44,300, supported by `item_17`, `item_18`. | 2 |
| **Total** | **All required facts supported** | **10/10** | **All required facts supported** | **10/10** |

Both explanations describe the two sheet roles and unused cancellation-fee data, and identify no refund or cancellation calculation. Terra explicitly declines to establish a reporting interval from the issue date. Luna describes January 2025 as the invoice's issue period, while explicitly declining to establish a broader reporting period; it does not claim a January transaction filter.

Neither explanation recovers the specification's JPY convention. Terra says that the workbook does not identify a currency; the evidence supports the narrower statement that currency was not identified in the exposed cell values. The tools cannot inspect currency number formats. Luna leaves currency unstated. Neither response establishes enforced input validation or timezone settings.

There are no incorrect required-fact or monetary claims in either final explanation. Terra's observation that cancellation fees equal 10% of listed prices is numerically correct, and it explicitly avoids promoting that coincidence into a general policy. Luna's extra reference to `Master!E5` is awkward but explicitly identifies it as unrelated to shipping and gives the correct shipping cell `D5`.

| Numeric result | Expected | Terra reported | Luna reported |
| --- | ---: | ---: | ---: |
| First line, `Invoice!H5` | 2,400 | 2,400 | 2,400 |
| Goods subtotal, `Invoice!H14` | 42,000 | 42,000 | 42,000 |
| Shipping total, `Invoice!H15` | 2,300 | 2,300 | 2,300 |
| Grand total, `Invoice!H16` | 44,300 | 44,300 | 44,300 |

Both traces also contain the correct remaining seven lines. The evaluator's independent arithmetic agrees with all captured invoice line amounts and totals. Both agents qualify their numeric outputs as saved caches and note that the tools do not recalculate them. Reading all cells and producing a correct explanation does not expose or prove every arithmetic operation performed internally by a model; the exhaustive comparisons remain evaluator work.

### Observable strategies and error recovery

Terra discovers sheets and names (`item_2`, `item_3`), then bounds both sheets with `A1:ZZ1000` searches (`item_4`, `item_5`). It reads the master before the invoice, alternating value and formula requests for `A1:H10` and `A11:H20` on each sheet (`item_6`–`item_13`). All requests are valid. Its rectangles extend beyond the populated data, returning 160 nonempty values in 320 requested value cells. It reaches the full invoice-to-master dependency path in its final explanation.

Luna discovers sheets/names and bounds them using smaller `A1:Z200` searches (`item_2`–`item_5`). It then requests four oversized reads:

| Trace item | Tool | Requested range | Result |
| --- | --- | --- | --- |
| `item_6` | Values | `Master!A1:F13` | Rejected: more than 10 rows |
| `item_7` | Formulas | `Master!A1:F13` | Rejected: more than 10 rows |
| `item_8` | Values | `Invoice!A1:H16` | Rejected: more than 10 rows |
| `item_9` | Formulas | `Invoice!A1:H16` | Rejected: more than 10 rows |

Each returns `excel_range must contain at most 10 rows and 10 columns`. These are real failed workbook requests, not setup failures, so all four count toward Luna's tool cost. They were issued as a batch before their failures were processed; the trace does not show four successive failed recovery attempts.

Luna's next commentary (`item_10`) acknowledges the limit and says it is splitting the sheets. It retries with valid blocks: `Master!A1:F10`, `A11:F13`, `Invoice!A1:H10`, and `A11:H16`, reading both values and formulas (`item_11`–`item_18`). Those eight calls succeed. It then explains the complete dependency path correctly. No supervisor hint or extra user prompt was supplied to assist recovery.

### Tool cost and interpretation

| Measure | Sol baseline | Terra | Luna |
| --- | ---: | ---: | ---: |
| Total MCP calls | 12 | 12 | 16 |
| Successful MCP calls | 12 | 12 | 12 |
| Failed MCP calls | 0 | 0 | 4 |
| Sheet / name / area calls | 1 / 1 / 2 | 1 / 1 / 2 | 1 / 1 / 2 |
| Value / formula calls, including failures | 4 / 4 | 4 / 4 | 6 / 6 |
| Exact repeated tool-and-argument requests | 0 | 0 | 0 |
| Proxy rejections | 0 | 0 | 0 |
| Value cells requested in successful reads, including blanks | 206 | 320 | 206 |
| Nonempty value cells returned | 160 | 160 | 160 |
| Formula cells returned | 51 | 51 | 51 |
| CLI input tokens | 68,556 | 56,023 | 63,772 |
| Cached input tokens, included above | 44,288 | 40,960 | 40,960 |
| Cache-write input tokens | 0 | 0 | 0 |
| CLI output tokens | 1,712 | 957 | 1,435 |
| Separately reported reasoning-output counter | 115 | 0 | 0 |
| Elapsed seconds | 54.418 | 32.577 | 37.973 |

Zero reasoning-output counters are recorded as emitted and do not establish that a model performed no reasoning; medium effort was explicitly requested for all three. No cash cost is inferred from the token counts. Luna's successful value-cell count excludes its rejected oversized requests, which returned error messages rather than cell data.

There were no additional setup failures for Terra or Luna. Both sessions had the same nonfatal under-development-feature warning about `skip_host_skill_discovery` as Sol. Neither exhausted its call or time budget, and neither had an unresolved dependency at completion.

The observed difference is successful first-pass tool use by Sol/Terra versus an initial range-size error followed by autonomous recovery by Luna. All three ultimately explained LEVEL_1 correctly. The elapsed-time and token differences are measurements from these individual runs, with different response lengths and potentially different backend/cache conditions. They should not be treated as stable performance rankings without repeated trials.
