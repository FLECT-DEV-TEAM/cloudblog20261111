# LEVEL_2 exploration evaluation

All three Codex runs explained the main receipts calculation and reported the correct **59,051,650** net receipts in `Summary!H18`. Sol covered all six evaluation criteria. Terra and Luna omitted an explicit explanation of transaction-row counts versus unit quantities. Luna additionally left the actual source-table extent unresolved and had repeated tool-use errors.

| Model | Required-fact score | MCP calls | Failed calls | Exact repeat requests | Source rows read | Elapsed seconds |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `gpt-6.1-sol` | 12/12 | 36 | 0 | 0 | 59 | 74.635 |
| `gpt-6.1-terra` | 11/12 | 32 | 0 | 0 | 19 | 45.311 |
| `gpt-6-luna` | 11/12 | 57 | 8 | 15 | 99 | 80.513 |

Source-row counts are distinct transaction rows returned by successful value reads, excluding the header. These are one-run observations per model, not stable model rankings. The scores assess explanations; tool failures, repetitions, and additional discovery gaps are reported separately.

## Conditions and artifacts

The experiment retained the LEVEL_1 policy: fresh ephemeral sessions, **medium** reasoning effort, at most **100 MCP calls**, a **600-second** supervisor time limit, and no additional token cap beyond the model's context/default output limits. Runs were sequential, in Sol/Terra/Luna order. No follow-up hints or retries of a completed experiment were supplied.

Each exploring process received this exact task, followed by the absolute path to `LEVEL_2/excel_level_2.xlsx`:

> Explore this workbook using the supplied Excel MCP tools. Explain its purpose, sheet roles, inputs, reporting period if any, and how its final monetary total is calculated. Trace representative formulas to their sources and cite sheet/cell references. Describe refund, cancellation, pricing, and exception rules where present. Distinguish verified facts from assumptions and identify anything you could not establish.

The experiment's developer constraints were unchanged from LEVEL_1. The model sessions did not receive the specification, expected results, canonical records, validator output, prior run results, rubric, or this evaluation. The supervising agent read these evaluator-only materials separately.

Shell execution, other apps/plugins, browsers, computer/image access, memory, and delegation were disabled. User configuration/rules, project-document loading, and automatic skills were disabled. Each process had an empty temporary working directory. A guarded stdio proxy forwarded only the five supplied tools to the unchanged `mcp.py`, and allowed only the assigned workbook path. Code Mode's host remained enabled because the CLI requires it for tool routing. The raw traces record only supplied MCP calls for workbook access.

| Field | Sol | Terra | Luna |
| --- | --- | --- | --- |
| Exact model requested | `gpt-6.1-sol` | `gpt-6.1-terra` | `gpt-6-luna` |
| Start, JST on 2026-11-11 | 11:58:16.837 | 11:59:40.434 | 12:01:21.741 |
| End, JST on 2026-11-11 | 11:59:31.581 | 12:00:25.842 | 12:02:42.443 |
| Session | `01a1c88a-b95d-7bb1-85df-7f1d1a65103d` | `01a1c88b-ffdf-7b01-a639-e8bbe95803be` | `01a1c88d-8bf5-7b71-8ce3-133436f66ccf` |
| Termination | Normal, exit 0 | Normal, exit 0 | Normal, exit 0 |

All runs used Codex CLI **0.167.1**, Python **3.14.10**, MCP SDK **1.30.0**, and openpyxl **3.1.5**. Elapsed durations use the supervisor's monotonic clock and include startup and final export. Model identifiers and reasoning effort are client selections; the events do not expose an immutable served backend snapshot.

| Artifact | Sol | Terra | Luna |
| --- | --- | --- | --- |
| Unedited final explanation | [Markdown](exploration_codex_gpt_6_1_sol01.md) | [Markdown](exploration_codex_gpt_6_1_terra01.md) | [Markdown](exploration_codex_gpt_6_luna01.md) |
| Raw CLI event stream | [JSONL](exploration_codex_gpt_6_1_sol01.jsonl) | [JSONL](exploration_codex_gpt_6_1_terra01.jsonl) | [JSONL](exploration_codex_gpt_6_luna01.jsonl) |
| Timestamped MCP requests/results | [Audit](exploration_codex_gpt_6_1_sol01.mcp.jsonl) | [Audit](exploration_codex_gpt_6_1_terra01.mcp.jsonl) | [Audit](exploration_codex_gpt_6_luna01.mcp.jsonl) |
| Exact prompt, command, versions, budgets, hashes | [Metadata](exploration_codex_gpt_6_1_sol01.metadata.json) | [Metadata](exploration_codex_gpt_6_1_terra01.metadata.json) | [Metadata](exploration_codex_gpt_6_luna01.metadata.json) |
| Derived metrics and numeric checks | [Metrics](exploration_codex_gpt_6_1_sol01.metrics.json) | [Metrics](exploration_codex_gpt_6_1_terra01.metrics.json) | [Metrics](exploration_codex_gpt_6_luna01.metrics.json) |
| CLI diagnostics | [Log](exploration_codex_gpt_6_1_sol01.stderr.log) | [Log](exploration_codex_gpt_6_1_terra01.stderr.log) | [Log](exploration_codex_gpt_6_luna01.stderr.log) |

## Independent reference and workbook integrity

The command `.venv/bin/python LEVEL_2/validation.py --require-caches` passed, checking all **20,000** source records, aggregation/summary formulas and cached results against independent policy calculations. It found **6,964** in-period rows and **zero missing formula caches**. Its captured result is saved in [exploration_validation.json](exploration_validation.json). This is evaluator verification, not work performed by the exploring models.

Every run recorded the same workbook SHA-256 before and after:

```text
7848801374f4d0dabc5474cac1f9ac4597dada4ea4cc2f4c21bb0ca86f0e2e9a
```

This matches the recalculated LEVEL_2 hash in [recalculation_manifest.json](../recalculation_manifest.json), which records Microsoft Excel **16.114.4**. The [expected-results file](expected_results.json) retains the earlier generation hash and pending-recalculation metadata; the later manifest and strict validation establish the experiment input's provenance. `mcp.py` and `tools.py` were unchanged throughout all runs.

The [offline checker](../experiments/evaluate_level2_trace.py) reconstructs values/formulas from the recorded MCP outputs. It independently computes policy results from saved canonical records using the evaluator's policy implementation and confirms agreement with the expected-results file. It compares only captured observations, reports missing evidence separately, and calculates trace costs. It does not open the workbook or make MCP/model calls. Reproduce a check with, for example:

```bash
.venv/bin/python experiments/evaluate_level2_trace.py --run-id exploration_codex_gpt_6_1_sol01
```

The active [runner](../experiments/run_codex_exploration.py) now accepts `--level 2` as well as `--model`. The prior LEVEL_1 runner used for Terra/Luna is preserved in [run_codex_exploration_level1.py](../experiments/run_codex_exploration_level1.py), keeping those recorded harness hashes reproducible. The LEVEL_1 artifacts and report were not altered by these tests.

## Required-fact evaluation

The [LEVEL_2 specification and experiment protocol](../excel_book_levels_specification.md) define six required fact groups. The [rubric](../experiments/level2_rubric.json) expands their expectations, including the intended challenge of distinguishing unit quantities from transaction-row counts. Scores are **0** missing/incorrect, **1** partially correct/unsupported, and **2** correct with relevant workbook evidence. The maximum is **12**, rather than LEVEL_1's five-criterion maximum of 10.

| Criterion | Sol | Terra | Luna |
| --- | --- | --- | --- |
| All six sheet roles | **2** — Describes Master, source transactions, all three status aggregations, controls, summary and reconciliation. | **2** — Identifies all six and explains their respective input/rollup/output roles. | **2** — Correctly describes all six roles. The unresolved source extent is discussed separately below. |
| Reporting boundaries | **2** — Cites `Summary!B2:C2`, exact inclusive/exclusive dates and transaction-timestamp predicates. | **2** — Gives the same controls, dates and `>=`/`<` logic. | **2** — Cites `B2`, `C1:C2`, and the shared start-inclusive/end-exclusive filtering. |
| Final-state row semantics | **2** — Describes final-status order lines, cites the source notes, separates SOLD and REFUNDED original charges, and explicitly avoids claiming a full duplicate audit. | **2** — Describes final-state order lines and mutually exclusive outcomes, while distinguishing these modeled outcomes from unverified order history. | **2** — Identifies order-line final statuses from source notes and correctly separates sold versus refunded charges in the model. |
| Status quantities and row counts | **2** — Explains status-filtered quantity sums with P001 examples for all statuses; explicitly says `COUNTIFS` counts matching rows without deduplicating IDs. | **1** — Correctly explains all three quantity aggregations and mentions transaction counts, but does not explain counts as rows versus quantities as units or the `COUNTIFS` distinction. | **1** — Correctly explains all three quantity aggregations and names count columns, but leaves their row-count semantics and contrast with quantities unexplained. |
| Original charges, refunds and fees | **2** — Explains sold goods plus per-unit shipping, original charges on refunded lines, goods-only refunds/retained shipping, and cancellation fees with waived amounts excluded. | **2** — Explains all three treatments and links them to master rates and the status sheets. | **2** — Explains original refunded charges, goods-only refunds, retained shipping and quantity-based cancellation fees; excludes waived goods/shipping. |
| Net receipts and supporting references | **2** — Gives `H6=E6-F6+G6`, `H18=SUM(H6:H17)`, 59,051,650 and the full Summary-to-status-aggregation-to-master/source chain. | **2** — Gives the same formula, total and representative P001 dependency chain. | **2** — Gives the same formula, total and representative source references through Sales/Refund/Cancel/Master. |
| **Total** | **12/12** | **11/12** | **11/12** |

The row-count deductions concern omitted explanation, not an incorrect quantity or demonstrated misunderstanding. All three obtained the relevant `COUNTIFS` formulas, but reading them is separate from explaining them. No agent audited every transaction for unique IDs; the final-state scores recognize the intended model interpretation supported by source notes and formulas, not proof that every source row conforms.

Representative evidence: Sol reads the status values in `item_12`–`item_19`, their formulas in `item_22`–`item_29`, summary formulas in `item_25`/`item_30`, and source notes in `item_11`/`item_33`–`item_37`. Terra's corresponding status/summary evidence is in `item_14`–`item_29`, with source samples in `item_30`/`item_32`. Luna's successful status values/formulas are in `item_38`–`item_43` and `item_52`–`item_57`, with summary evidence in `item_44`/`item_45` and `item_58`/`item_59`.

## Numeric accuracy and additional limitations

All final monetary claims checked against the independent oracle are correct. The common headline values are:

| Result | Independent expected value | Sol | Terra | Luna |
| --- | ---: | ---: | ---: | ---: |
| P001 initial charges, `Summary!E6` | 1,920,000 | 1,920,000 | 1,920,000 | 1,920,000 |
| P001 refunds paid, `Summary!F6` | 393,000 | 393,000 | 393,000 | 393,000 |
| P001 cancellation fees, `Summary!G6` | 18,300 | 18,300 | 18,300 | 18,300 |
| P001 net receipts, `Summary!H6` | 1,545,300 | 1,545,300 | 1,545,300 | 1,545,300 |
| Net receipts total, `Summary!H18` | **59,051,650** | **59,051,650** | **59,051,650** | **59,051,650** |

Sol and Terra also explicitly report the correct grand-total bridge: **74,028,800 initial charges − 15,756,000 refunds + 778,850 fees = 59,051,650**. Luna states this bridge symbolically and gives its correct P001 example. All three read the grand-total components. No monetary error was found in the final explanations or in the captured values compared by the checker.

The independent source calculation also gives **14,551 sold units, 4,316 refunded units, and 2,073 canceled units**, with **4,818 / 1,457 / 689 transaction rows** respectively. Thus 20,940 in-period units belong to 6,964 rows. These are evaluator totals, not numbers independently recounted by the agents. The distinction illustrates the omitted explanation in the Terra/Luna responses.

Every run returned all **357 calculation-sheet formulas**, and they matched the expected formulas. All checked captured values matched: **880** for Sol, **640** for Terra, and **1,120** for Luna. These counts include headers and sampled source data; they are not raw-transaction coverage counts. Missing observations outside the samples are not treated as wrong values.

All explanations correctly warn that the tools return saved formula caches without recalculation. Sol and Terra identify the zero-shipping products and explain why refunded shipping can be zero. All three explain the reconciliation equivalence. Sol and Terra explicitly state that zero reconciliation differences establish internal arithmetic consistency, not complete/valid source data. No model verified the exact reporting-boundary source rows by reading the evaluator's known boundary addresses.

None establishes the specification's **JPY** or Japan-time convention from the cell data. Terra's categorical currency statement, and Luna's similar wording, should be narrowed to currency not being identified in the exposed values; these tools do not inspect number formats. No agent establishes full-line refund timing after shipment or cancellation timing before shipment as an externally documented business policy. The calculations implement the specified treatment, but worksheet formulas and final-state notes do not establish a real business's operational history.

Luna's extra structural gap is material: it leaves the full Transactions extent unresolved. Its search was only `A1:Z100`, so `get_value_area` returned `A1:F100` **within that search window**. The actual table has 20,000 records through row 20001. Luna notes that formulas reach row 20001 and explicitly says it could not establish whether additional records exist, rather than claiming a complete source audit. That uncertainty is honest, but a wider area search or late-row read could have resolved it. The required-fact score does not separately allocate a point for population size, so this gap must remain visible alongside the score.

## Observable exploration behavior

**Sol** discovers sheets/names, bounds all six sheets using `A1:AZ100000`, reads the small master/status/summary tables, then reads their formula blocks. It subsequently expands its transaction sample to rows **2–50 and 19992–20001**. Those 59 rows include all three statuses and December/January/February timestamps. It successfully establishes the large table's extent without reading it all. The final explanation traces the net-receipts total down through status quantities and master rates, and includes a raw P001 refund example at `Transactions!A12:F12`.

**Terra** discovers names/sheets and bounds them using `A1:AZ50000`. It alternates values and formulas across the small tables, then samples transaction rows **2–10 and 19992–20001**. The 19 source rows include all three statuses and all three months. Its complete small-table reads include all 357 calculation formulas; it does not need 2,000 source-data windows to explain the repeated aggregation pattern.

**Luna** bounds sheets using `A1:Z100` and reads source rows **2–100** in ten windows. It later repeats those same ten source windows, the two master windows, and two summary windows. It never extends the source sample beyond row 100. Consequently, it reads more unique source rows than either other model while failing to resolve the full table extent.

Luna also makes eight oversized requests:

| Trace items | Tool and range | Result |
| --- | --- | --- |
| `item_21` | Values, `Refund!A1:I13` | Rejected: more than 10 rows |
| `item_22` | Values, `Sales!A1:G13` | Rejected: more than 10 rows |
| `item_23` | Values, `Cancel!A1:G13` | Rejected: more than 10 rows |
| `item_47` | Formulas, `Sales!A1:G13` | Rejected: more than 10 rows |
| `item_48` | Formulas, `Refund!A1:I13` | Rejected: more than 10 rows |
| `item_49` | Formulas, `Cancel!A1:G13` | Rejected: more than 10 rows |
| `item_50`, `item_51` | Formulas, `Summary!A1:J18`, twice | Both rejected: more than 10 rows |

The tools return `excel_range must contain at most 10 rows and 10 columns`. Luna recovers by splitting the ranges, but later repeats the same size mistake when switching from values to formulas. Its **15 exact repeat requests** comprise 14 repeated successful value reads and one repeated rejected Summary formula request. These counts describe visible requests, not inferred reasons for making them.

Sol and Terra have no failed or exact repeated requests. Luna shows observable difficulty with tool constraints and efficient extent discovery; it nevertheless reaches and explains the correct calculation before either budget is exhausted. Its tool-use problems should not be presented as an incorrect net-receipts explanation.

## Tool cost and verification scope

| Measure | Sol | Terra | Luna |
| --- | ---: | ---: | ---: |
| Total supplied MCP calls | 36 | 32 | 57 |
| Successful / failed | 36 / 0 | 32 / 0 | 49 / 8 |
| Sheet / name / area calls | 1 / 1 / 6 | 1 / 1 / 6 | 1 / 0 / 6 |
| Value / formula calls, including failures | 16 / 12 | 12 / 12 | 37 / 13 |
| Exact repeated requests | 0 | 0 | 15 |
| Proxy rejections | 0 | 0 | 0 |
| Unique source rows read | 59 | 19 | 99 |
| Percentage of 20,000 source rows | 0.295% | 0.095% | 0.495% |
| Value cells requested in successful calls, including blanks/repeats | 1,100 | 1,120 | 2,015 |
| Unique nonempty value cells returned | 881 | 641 | 1,121 |
| Unique formula cells returned | 357 | 357 | 357 |
| CLI input tokens | 142,381 | 70,004 | 154,831 |
| Cached input tokens, included above | 111,744 | 41,472 | 118,784 |
| Cache-write input tokens | 0 | 0 | 0 |
| CLI output tokens | 2,353 | 1,555 | 2,654 |
| Separately reported reasoning-output counter | 364 | 27 | 199 |
| Elapsed seconds | 74.635 | 45.311 | 80.513 |

Input-token totals include context across model requests. Counts are emitted CLI counters, not a cash bill or an inference about private reasoning. Same-range value/formula reads serve different purposes and are not counted as exact repeats. Rejected reads returned no cell data; their calls count toward costs, while their hypothetical cell areas do not count as returned values.

There were no setup failures or supervisor interventions in these LEVEL_2 runs. All sessions emitted the same nonfatal warning that `skip_host_skill_discovery` is an under-development feature. No run exhausted its 100-call/600-second limits. These traces support the intended scale strategy: inspect the small calculation tables and sample the large source. They do not demonstrate independent agent verification of all 20,000 rows, boundary-case coverage, or global data integrity. Those stronger checks belong to the separate evaluator validation recorded above.

The results support successful recovery of the main LEVEL_2 business logic, with an explanation omission for Terra/Luna and additional tool-use/discovery weaknesses for Luna. Repeated trials are needed to assess how consistently each model behaves this way.
