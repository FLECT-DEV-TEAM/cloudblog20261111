# Excel Workbook Levels: Specification for an AI Exploration Experiment

## Purpose

Create three synthetic Excel workbooks that let us observe how an AI agent discovers worksheet structure, samples data, follows formulas, and explains business logic through partial reads.

The levels separate three sources of difficulty:

| Level | Main challenge | Expected exploration |
| --- | --- | --- |
| LEVEL_1 | A small calculation spanning two sheets | Read the invoice and follow direct master lookups. |
| LEVEL_2 | A large transaction table with simple aggregation | Inspect representative records and infer repeated `SUMIFS` patterns without reading every record. |
| LEVEL_3 | Distributed rules and several layers of dependencies | Trace summaries through calculations, mappings, effective dates, and exceptions. |

The hypothesis is that agents will explain LEVEL_1 and LEVEL_2 more reliably than LEVEL_3. This is an experimental hypothesis, not a required outcome. A correct LEVEL_3 explanation is a successful observation, too.

This document specifies workbooks to be created later; it does not report measured results.

## Common conventions

- File names: `excel_level_1.xlsx`, `excel_level_2.xlsx`, and `excel_level_3.xlsx`.
- Use synthetic data only, English sheet names and headers, and integer JPY amounts. Exclude taxes, exchange rates, inventory costs, and accounting recognition rules.
- “Purchase amount” means **quantity of units**, not a monetary amount. Quantities are positive integers. Shipping and cancellation fees are charged **per unit**, not per order or per transaction row.
- Store datetimes as Excel datetime values, interpreted in Japan local time. Use a reporting interval with an inclusive start and an exclusive end.
- Use ordinary worksheet formulas, stored values, and workbook-defined names. Do not depend on macros, external workbooks, Power Query, charts, comments, or formatting to convey required business rules.
- Use visible headers, consistent number formats, and freeze panes. Formatting is helpful for a human but must not be the only source of meaning.
- All required dependencies must remain inside the workbook. Formula ranges must cover every intended data row and must exclude totals rows.
- Formula examples below use English function names and comma separators. Copy-down instructions apply to the entire specified data block.

### Tool environment

The experiment uses the read-only MCP server in `mcp.py`, with these tools from `tools.py`:

| Tool | Available information and constraints |
| --- | --- |
| `get_sheet_names` | All sheet names, including hidden sheets. |
| `get_named_regions` | Defined names, their scopes, references, and hidden flags. Names are not evaluated. |
| `get_value_area` | Smallest occupied rectangle within a supplied finite A1 search rectangle. It does not enumerate separate tables or automatically discover the entire used range. |
| `get_cell_values_csv` | Saved values in a rectangle of at most **10 rows × 10 columns**. No headers or coordinates are added by the tool. Formula results come from saved caches. |
| `get_cell_formulas` | Formula text keyed by cell address, in a rectangle of at most **10 rows × 10 columns**. Constants are omitted and formulas are not evaluated. |

The 10 × 10 limit concerns the dimensions of each read, not the position of the range: `K100:T109` is valid. `get_value_area` has no corresponding 10 × 10 restriction, but returns only a bounding rectangle. Large bounding searches can still cost time.

Before an experiment, recalculate each workbook in a spreadsheet engine that supports its formulas and save current cached results. Verify the saved caches through the MCP tools. Writing formulas with a library alone does not establish that their cached results exist. Freeze the workbook during a run; the server caches opened workbooks, so restart it when replacing files.

### Shared product master

All three levels use the same `Master!A1:F13` table. Row 1 contains these headers:

| Column | Header | Meaning |
| --- | --- | --- |
| A | ProductCode | Unique canonical product identifier. |
| B | ProductName | Human-readable synthetic name. |
| C | UnitPrice | Baseline price per unit. |
| D | ShippingFeePerUnit | Shipping fee per unit. |
| E | CancellationFeePerUnit | Fee per unit canceled before shipment. |
| F | Category | Product category. |

Use this fixed data so that small calculations are easy to verify:

| ProductCode | ProductName | UnitPrice | ShippingFeePerUnit | CancellationFeePerUnit | Category |
| --- | --- | ---: | ---: | ---: | --- |
| P001 | Product 01 | 1000 | 200 | 100 | A |
| P002 | Product 02 | 1500 | 300 | 150 | A |
| P003 | Product 03 | 2000 | 200 | 200 | A |
| P004 | Product 04 | 2500 | 0 | 250 | A |
| P005 | Product 05 | 3000 | 300 | 300 | B |
| P006 | Product 06 | 3500 | 200 | 350 | B |
| P007 | Product 07 | 4000 | 300 | 400 | B |
| P008 | Product 08 | 4500 | 0 | 450 | B |
| P009 | Product 09 | 5000 | 200 | 500 | C |
| P010 | Product 10 | 5500 | 300 | 550 | C |
| P011 | Product 11 | 6000 | 200 | 600 | C |
| P012 | Product 12 | 6500 | 0 | 650 | C |

## LEVEL_1: A Simple Invoice Sheet

Two sheets: **Master** and **Invoice**.

### Master

Use the shared product master without changes. The cancellation fee is present for consistency with the later levels but does not participate in this invoice.

### Invoice

Create one invoice with eight populated purchase lines. Store an invoice identifier in `B2` and its issue date in `F2`.

Headers are in `A4:H4`; purchase lines occupy rows 5–12:

| Column | Header | Source or calculation |
| --- | --- | --- |
| A | ProductCode | Input; a valid master code. |
| B | ProductName | Exact-match master lookup. |
| C | Quantity | Input; positive integer. |
| D | UnitPrice | Exact-match master lookup. |
| E | ShippingFeePerUnit | Exact-match master lookup. |
| F | GoodsSubtotal | Quantity × unit price. |
| G | ShippingSubtotal | Quantity × shipping fee per unit. |
| H | LineTotal | Goods subtotal + shipping subtotal. |

Representative formulas in row 5, copied down to row 12:

```excel
B5 =VLOOKUP(A5,Master!$A$2:$F$13,2,FALSE)
D5 =VLOOKUP(A5,Master!$A$2:$F$13,3,FALSE)
E5 =VLOOKUP(A5,Master!$A$2:$F$13,4,FALSE)
F5 =C5*D5
G5 =C5*E5
H5 =F5+G5
```

Put labels in `A14:A16` and the following totals in column H:

```excel
H14 =SUM(F5:F12)
H15 =SUM(G5:G12)
H16 =SUM(H5:H12)
```

The labels are `Goods subtotal`, `Shipping total`, and `Grand total`, respectively. `H16` must also equal `H14+H15`.

Use product codes P001–P008 in order, with quantities `2, 1, 3, 2, 1, 2, 1, 3`. For example, the first two line totals are JPY 2,400 and JPY 1,800. The invoice goods subtotal is JPY 42,000, shipping total is JPY 2,300, and grand total is JPY 44,300.

### Intended exploration challenge

An agent should identify the inputs, follow an exact-match lookup into `Master`, explain the per-unit shipping rule, and cite the grand-total cell. A single 10 × 10 read cannot contain the entire invoice, but only a few targeted reads should be needed. No defined names or hidden sheets are required.

## LEVEL_2: A Transaction Summary

Six sheets: **Master**, **Transactions**, **Sales**, **Refund**, **Cancel**, and **Summary**.

### Business rules

These are explicit policies for this synthetic business; refund and cancellation policies vary between real businesses.

Each transaction row represents **one order line in its final state**, not a sequence of events. A refunded order line appears once as `REFUNDED`; it does not also have a `SOLD` row. All quantities and raw amounts are positive. Assign the row to a reporting period using its `TransactionDateTime`, not a separate refund date.

| Final status | Initial charge | Refund paid | Cancellation fee | Net receipts |
| --- | --- | --- | --- | --- |
| SOLD | Goods + shipping | 0 | 0 | Goods + shipping |
| REFUNDED | Goods + shipping | Goods only | 0 | Shipping retained |
| CANCELED | 0 | 0 | Quantity × cancellation fee per unit | Cancellation fee |

Assume full-line refunds after shipment and cancellations before shipment. Do not model partial returns. “Net receipts” is the experiment's cash calculation, not profit or a claim about financial reporting treatment.

### Master

Use the same table and values as LEVEL_1. Prices and fees are constant across the dataset in LEVEL_2.

### Transactions

Create **20,000 rows** in `A2:F20001`, with headers in row 1:

| Column | Header | Requirement |
| --- | --- | --- |
| A | TransactionID | Unique: TX000001 through TX020000. |
| B | TransactionDateTime | Excel datetime; includes records before, during, and after January 2025. |
| C | ProductCode | One of the 12 master codes. |
| D | Quantity | Integer from 1 to 5. |
| E | Status | SOLD, REFUNDED, or CANCELED. |
| F | Note | Optional synthetic operational text; not an input to calculations. |

Use a documented pseudorandom seed, e.g. `504`, and record the generator version. Target a 70% / 20% / 10% status distribution. Cover December 2024 through February 2025; include records exactly at both reporting boundaries. Ensure every product has all three statuses within January, and include quantities greater than one and zero-shipping products among refunded records. Shuffle row order deterministically so head-only sampling is insufficient for time coverage.

The table's size is intended to discourage reading the complete dataset. Whether it exceeds a particular model's context depends on serialization and model limits; the experiment must record those limits rather than assume them.

### Reporting controls

In `Summary`, put `PeriodStart` in `A2`, its value `2025-01-01 00:00:00` in `B2`, and the exclusive `PeriodEnd` value `2025-02-01 00:00:00` in `C2`. Put the label `PeriodEnd (exclusive)` in `C1`.

All aggregation sheets must reference these same controls. Include period criteria in every transaction-level aggregation:

```excel
Transactions!$B$2:$B$20001,">="&Summary!$B$2,
Transactions!$B$2:$B$20001,"<"&Summary!$C$2
```

### Sales

Use `A1:G13`, one row per product, ordered P001–P012:

`ProductCode | ProductName | SoldQuantity | GoodsAmount | ShippingAmount | InitialCharge | TransactionCount`

Representative formulas for the product code in `A2`:

```excel
C2 =SUMIFS(Transactions!$D$2:$D$20001,Transactions!$C$2:$C$20001,$A2,Transactions!$E$2:$E$20001,"SOLD",Transactions!$B$2:$B$20001,">="&Summary!$B$2,Transactions!$B$2:$B$20001,"<"&Summary!$C$2)
D2 =C2*VLOOKUP(A2,Master!$A$2:$F$13,3,FALSE)
E2 =C2*VLOOKUP(A2,Master!$A$2:$F$13,4,FALSE)
F2 =D2+E2
```

Use an exact-match lookup for `ProductName`. Use `COUNTIFS` with the same product, status, and date criteria for `TransactionCount`; this counts rows rather than units.

### Refund

Use `A1:I13`:

`ProductCode | ProductName | RefundedQuantity | OriginalGoods | OriginalShipping | InitialCharge | RefundPaid | RetainedShipping | TransactionCount`

- Column C uses the same `SUMIFS` pattern as `Sales`, with status `REFUNDED`.
- D = C × master unit price; E = C × master shipping fee per unit.
- F = D + E; G = D; H = F − G.
- Column I counts matching transaction rows with `COUNTIFS`.

This sheet explicitly preserves the original charge so that `Summary` can subtract refunds without counting refunded rows as sales a second time.

### Cancel

Use `A1:G13`:

`ProductCode | ProductName | CanceledQuantity | WaivedGoods | WaivedShipping | CancellationFees | TransactionCount`

- Column C uses `SUMIFS` with status `CANCELED`.
- D = C × master unit price; E = C × master shipping fee per unit. These are informational waived amounts and never enter initial charges.
- F = C × master cancellation fee per unit.
- Column G counts matching transaction rows with `COUNTIFS`.

### Summary

Use headers in `A5:J5`, product rows 6–17, and totals in row 18:

`ProductCode | SoldQuantity | RefundedQuantity | CanceledQuantity | InitialCharges | RefundsPaid | CancellationFees | NetReceipts | RetainedRefundShipping | ReconciliationDifference`

For product P001 in row 6, linking to row 2 on the aggregation sheets:

```excel
B6 =Sales!C2
C6 =Refund!C2
D6 =Cancel!C2
E6 =Sales!F2+Refund!F2
F6 =Refund!G2
G6 =Cancel!F2
H6 =E6-F6+G6
I6 =Refund!H2
J6 =H6-(Sales!F2+Refund!H2+Cancel!F2)
```

Copy the pattern down for all products. Sum each numeric column in row 18. Every reconciliation difference must be zero.

For a small independent example using P001, suppose a reporting period contains 2 sold units, 1 refunded unit, and 3 canceled units. Initial charges are JPY 3,600, refunds paid are JPY 1,000, cancellation fees are JPY 300, and net receipts are JPY 2,900. This example is a policy check, not the expected total of the generated 20,000-row dataset.

### Intended exploration challenge

An agent should discover the reporting controls, distinguish row counts from quantities, explain all three final statuses, and follow the totals through the aggregation sheets. It should infer the regular calculation pattern from representative formulas and seek counterexamples rather than copying all 20,000 rows into context. No defined names or hidden sheets are required.

## LEVEL_3: A More Complex Transaction Summary

Extend LEVEL_2 to **11 sheets**: retain its six sheets and add **StatusMap**, **ProductAlias**, **PriceHistory**, **Rules**, and **Calc**.

Preserve the same 20,000 transaction identities, timestamps, canonical products, quantities, and final statuses. Introduce alternate representations and policy changes so the extra difficulty comes mainly from tracing logic. LEVEL_3 totals may differ because its business rules differ.

### Transactions: source-system representations

Use `A1:H20001` with these columns:

`TransactionID | TransactionDateTime | RawProductCode | Quantity | RawStatus | Channel | SourceSystem | Note`

Replace some product and status values with the aliases defined below. Assign `DIRECT`, `PARTNER`, or `VIP` channels deterministically, with all channels represented across products and statuses. `SourceSystem` records which synthetic export representation was used; `Note` remains informational. Retain the original canonical dataset outside the workbook as part of the evaluator's reference data.

### StatusMap and ProductAlias

- `StatusMap!A1:B10`: `RawStatus | CanonicalStatus`. Include identity mappings for the three canonical statuses, plus `S` and `PAID` → SOLD, `R` and `RETURN` → REFUNDED, and `C` and `VOID` → CANCELED.
- `ProductAlias!A1:B37`: `RawProductCode | ProductCode`. For each product index 001–012, map `P001`, `OLD001`, and `SKU001` to P001, and similarly for the other products.
- Require unique raw lookup keys. Every source row must resolve exactly once. Unknown keys must produce a visible error rather than silently become zero or a default product.

### PriceHistory: effective-dated master values

Use `A1:H25`:

`ProductCode | EffectiveFrom | EffectiveToExclusive | UnitPrice | ShippingFeePerUnit | CancellationFeePerUnit | Version | PriceKey`

Create two rows per product:

1. `V1`: from 2024-12-01 through, but excluding, 2025-01-16. Use the shared master values.
2. `V2`: from 2025-01-16 through, but excluding, 2025-03-01. Increase unit price by 10%; keep shipping and cancellation fees unchanged.

`PriceKey` is `ProductCode & "|" & Version`, e.g. `P001|V2`. For each product, intervals must neither overlap nor leave a gap in the transaction date coverage. Use the order's transaction timestamp for price selection, including refunded and canceled rows. `Master` is now a baseline catalog; it is not the authoritative source of transaction prices.

### Rules: channel policies and approved adjustments

The primary policy table occupies `Rules!A1:E10`:

`RuleKey | DiscountRate | RefundShipping | CancellationMultiplier | Description`

There is one row for every combination of the three channels and three canonical statuses. Keys are `Channel & "|" & CanonicalStatus`, e.g. `VIP|REFUNDED`. Apply these values to all three status rows for each channel:

| Channel | DiscountRate | RefundShipping | CancellationMultiplier |
| --- | ---: | ---: | ---: |
| DIRECT | 0 | 0 | 1 |
| PARTNER | 0.05 | 0 | 1.5 |
| VIP | 0.10 | 1 | 0 |

Discounts apply only to goods. Refunds return the discounted goods amount actually charged. A `RefundShipping` value of 1 additionally returns shipping for a refunded line. Cancellation multipliers affect only the cancellation fee, never the goods amount. A VIP cancellation therefore has no fee.

Place a separate adjustment table in `Rules!K1:N21`:

`TransactionID | AdjustmentAmount | Reason | ApprovalReference`

Create 20 unique entries for transaction numbers 997 × k, where k = 1…20. Alternate +500 for odd k and −500 for even k. Use explicit synthetic reasons and approval references. Absence from this table means zero adjustment; duplicate IDs are invalid. Apply the signed adjustment once, regardless of final status, and only include it in a report when the transaction timestamp is in that report's period. Ensure at least one adjusted transaction falls in January when generating the shared dataset.

This second table deliberately lies outside a first `A1:J10` read. Its existence must be discoverable through defined names and downstream formulas.

### Calc: normalized, enriched transaction calculations

Create one calculation row per transaction in `A2:V20001`, with headers in row 1:

| Columns | Headers |
| --- | --- |
| A–F | TransactionID, TransactionDateTime, ProductCode, Quantity, CanonicalStatus, Channel |
| G–L | Category, PriceKey, UnitPrice, ShippingFeePerUnit, CancellationFeePerUnit, RuleKey |
| M–R | DiscountRate, RefundShipping, CancellationMultiplier, GoodsAmount, ShippingAmount, InitialCharge |
| S–V | RefundPaid, CancellationFees, Adjustment, NetReceipts |

Link A, B, D, and F to the corresponding source row. Resolve C and E with exact-match `XLOOKUP`; look up G from `Master`. Use workbook-defined names for lookup arrays as specified below.

Representative row-2 formulas:

```excel
C2 =XLOOKUP(Transactions!C2,AliasKeys,AliasProducts,NA(),0)
E2 =XLOOKUP(Transactions!E2,RawStatuses,CanonicalStatuses,NA(),0)
H2 =LET(p,C2,t,B2,XLOOKUP(1,(PriceProducts=p)*(PriceFrom<=t)*(PriceTo>t),PriceKeys,NA(),0))
I2 =XLOOKUP(H2,PriceKeys,PriceValues,NA(),0)
J2 =XLOOKUP(H2,PriceKeys,PriceShipping,NA(),0)
K2 =XLOOKUP(H2,PriceKeys,PriceCancellation,NA(),0)
L2 =F2&"|"&E2
M2 =XLOOKUP(L2,RuleKeys,RuleDiscounts,NA(),0)
N2 =XLOOKUP(L2,RuleKeys,RuleRefundShipping,NA(),0)
O2 =XLOOKUP(L2,RuleKeys,RuleCancellationMultipliers,NA(),0)
P2 =ROUND(D2*I2*(1-M2),0)
Q2 =D2*J2
R2 =IF(E2="CANCELED",0,P2+Q2)
S2 =IF(E2="REFUNDED",P2+IF(N2=1,Q2,0),0)
T2 =IF(E2="CANCELED",ROUND(D2*K2*O2,0),0)
U2 =XLOOKUP(A2,AdjustmentIDs,AdjustmentAmounts,0,0)
V2 =R2-S2+T2+U2
```

Round the discounted goods amount and cancellation fee **per transaction row** to whole JPY before aggregating. Do not round only the final product total. The adjustment lookup's zero fallback is intentional; missing mappings elsewhere must remain errors. Validate the effective-date matches independently: `XLOOKUP` returning one match does not prove that intervals are unique.

Create workbook-scoped defined names for these exact ranges:

| Names | Corresponding ranges, in the same order |
| --- | --- |
| AliasKeys, AliasProducts | `ProductAlias!$A$2:$A$37`, `ProductAlias!$B$2:$B$37` |
| RawStatuses, CanonicalStatuses | `StatusMap!$A$2:$A$10`, `StatusMap!$B$2:$B$10` |
| PriceProducts, PriceFrom, PriceTo | `PriceHistory!$A$2:$A$25`, `PriceHistory!$B$2:$B$25`, `PriceHistory!$C$2:$C$25` |
| PriceValues, PriceShipping, PriceCancellation | `PriceHistory!$D$2:$D$25`, `PriceHistory!$E$2:$E$25`, `PriceHistory!$F$2:$F$25` |
| PriceKeys | `PriceHistory!$H$2:$H$25` |
| RuleKeys, RuleDiscounts, RuleRefundShipping, RuleCancellationMultipliers | `Rules!$A$2:$A$10`, `Rules!$B$2:$B$10`, `Rules!$C$2:$C$10`, `Rules!$D$2:$D$10` |
| AdjustmentIDs, AdjustmentAmounts | `Rules!$K$2:$K$21`, `Rules!$L$2:$L$21` |

### Sales, Refund, Cancel, and Summary

Keep the LEVEL_2 layouts and reporting controls where possible, but aggregate calculated **row amounts from `Calc`** using `SUMIFS`. Multiplying aggregate quantity by a single current price is no longer correct.

- `Sales`: filter `Calc!E:E` to SOLD; sum quantity D, goods P, shipping Q, and initial charge R, using bounded rows 2–20001.
- `Refund`: filter to REFUNDED; sum D, P, Q, R, and S into the corresponding columns. Retained shipping is original shipping minus the shipping portion refunded; under these rules it also equals initial charge minus refund paid.
- `Cancel`: filter to CANCELED; sum D, P, Q, and T. Waived goods now reflect the discounted amount. Canceled goods and shipping still do not enter initial charges.
- All three sheets continue counting transaction rows separately and applying the shared reporting interval to `Calc!B2:B20001`.
- Extend `Summary` with `K5 = Adjustments`, and sum `Calc!U2:U20001` by product and reporting interval into K6:K17. K18 is the column total.
- Change each summary net-receipts formula to `InitialCharges - RefundsPaid + CancellationFees + Adjustments`.
- Change each reconciliation formula to compare that result against a direct `SUMIFS` of `Calc!V2:V20001` for the same product and period. Sum differences in J18, and require every product difference to be zero.

An example dependency path is:

`Summary net receipts → Refund refund paid → Calc refund → discounted historical price + channel refund policy → raw product/status mappings`.

The signed adjustment is a separate dependency that must also be explained.

### Required edge cases and discoverability

- Include transactions immediately before and exactly at the 2025-01-16 price transition for the same product.
- Include both a VIP refund that returns shipping and a DIRECT refund that retains nonzero shipping.
- Include a VIP cancellation with zero fee and a PARTNER cancellation with a 1.5 multiplier.
- Include a PARTNER cancellation, e.g. P002, quantity 1: JPY 225. The prescribed cancellation fees produce whole-JPY results even after the multiplier; retain the explicit rounding formula for a consistent calculation rule.
- Include at least one discounted goods amount requiring rounding, e.g. P002 V2, quantity 1, PARTNER: JPY 1,567.5 rounds to JPY 1,568.
- Place qualifying examples beyond the first ten transaction rows; record their actual row addresses in the evaluator's reference file.
- Keep lookup and rule sheets visible in the primary experiment. A secondary variant may hide `StatusMap` and `ProductAlias`, but report this as a separate condition.
- Use `XLOOKUP`, `LET`, `IF`, `ROUND`, `SUMIFS`, composite keys, and defined names to introduce realistic dependencies. Do not introduce arbitrary formula obfuscation, deliberate calculation errors, or missing external files.

### Intended exploration challenge

An agent should move between multiple 10 × 10 windows, resolve names, follow the historical-price selection, recognize policy branches and rounding order, and discover the separate adjustment block. Reading only the displayed summary or the first formula pattern is insufficient for a complete explanation.

## Generation and validation requirements

Keep this specification and the independent expected-results files outside the agent's allowed experiment inputs.

1. Generate all workbooks reproducibly. Record the seed, generator version, workbook hashes, recalculation engine, and its version.
2. Produce a separate reference file containing sheet/range inventories, the policy definitions, expected product and grand totals, and representative dependency paths with cell addresses.
3. Compute expected totals in independent code from canonical source records and policy data, rather than copying the workbook's cached summary values. Validate quantities, row counts, charges, refunds, fees, adjustments, and net receipts.
4. Check unique IDs and lookup keys, valid statuses, complete price coverage, reporting boundaries, non-overlapping price intervals, and absence of formula errors or missing numeric caches.
5. Confirm that LEVEL_2 status quantities and counts reconcile to the in-period source records. Confirm all LEVEL_2 and LEVEL_3 product reconciliation differences are zero, and compare all totals against the independent reference.
6. As a generation check, neutralize LEVEL_3 policies in a separate temporary variant: use baseline prices for both versions, zero discounts, retain refunded shipping, use cancellation multiplier 1 for every channel, and remove adjustments. Its numeric report totals must equal LEVEL_2 after alias normalization. Do not use this variant as the main LEVEL_3 test.
7. Verify several valid 10 × 10 reads, including late transaction rows and `Rules!K1:N10`, through the actual MCP server. Confirm that larger value/formula requests are rejected.

## Experiment protocol and observations

Use the same task prompt for both agents and all levels, changing only the workbook path:

> Explore this workbook using the supplied Excel MCP tools. Explain its purpose, sheet roles, inputs, reporting period if any, and how its final monetary total is calculated. Trace representative formulas to their sources and cite sheet/cell references. Describe refund, cancellation, pricing, and exception rules where present. Distinguish verified facts from assumptions and identify anything you could not establish.

Use a fresh session for each run. For the primary experiment, permit workbook access only through the five supplied MCP tools; disable alternative filesystem, Python, shell, or spreadsheet access to the workbook and evaluator files. Otherwise, an agent could bypass the partial-read constraint. Record exact model identifiers, agent versions, prompts, reasoning settings where exposed, tool-call and token budgets, elapsed time, and termination conditions. Choose budgets before observing outcomes and use the same budget policy across agents. Repeat runs to assess variability.

Record observable actions and explanations; do not infer access to an agent's private reasoning. Useful observations include discovery order, top-down versus source-first exploration, value/formula read balance, sampled ranges, dependency depth reached, repeated reads, recovery from tool errors, and explicit uncertainty.

Evaluate independently against these required facts:

| Level | Required facts |
| --- | --- |
| LEVEL_1 | Inputs and master lookups; quantity-based shipping; subtotal and grand-total formulas and values. |
| LEVEL_2 | All six sheet roles; reporting boundaries; final-state row semantics; status-specific quantities; original-charge/refund/fee treatment; net receipts and supporting references. |
| LEVEL_3 | LEVEL_2 concepts plus alias normalization, effective-date prices, channel discounts, shipping refund branches, cancellation multipliers, per-row rounding, signed adjustments, and the complete aggregate dependency path. |

Score each required fact as 0 (missing or incorrect), 1 (partially correct or unsupported), or 2 (correct with relevant workbook evidence). Report factual coverage, incorrect claims, numeric accuracy, and tool cost separately. Distinguish discovering a formula from independently verifying all source rows; sampling alone does not prove global data integrity.

Describe an agent as struggling only when the trace and evaluation support that description—for example, missed adjustment rules, repeated unresolved dependencies, incorrect totals, or budget exhaustion. Report successful exploration and uncertainty just as explicitly. The blog should separate these observations from the initial expectation about LEVEL_3.
