# Build guide: SOFIPO and fintech savings in Power BI (web)

You build this yourself in the Power BI Service in the browser, with your Tec de Monterrey Microsoft account. Nothing here needs Power BI Desktop. Plan on 3 to 4 hours spread over a couple of sessions.

**The question the report answers:** I have 200,000 MXN. How do rate changes affect it over 1, 3, 6 and 12 months, and how should I split it across fintech and SOFIPO products, given the ISR withholding, the SOFIPO tax exemption and the deposit protection limits?

Learning project only, not financial advice. See the disclaimer in README.md.

---

## 0. Know the rules the model uses

| Rule | Value (2026) | Where it is |
|---|---|---|
| ISR provisional withholding | 0.90% per year, charged on the **capital**, not on the interest | `tax_params` |
| 5-UMA exemption (LISR art. 93 fr. XX) | Interest is exempt on an average daily balance up to 5 annual UMAs = **213,973.20 MXN**, counted across all eligible balances. Above that, only the excess is withheld. Eligible: every SOFIPO product, and bank **on-demand savings** accounts when the bank applies it (Nu confirmed; other banks marked `unknown` and taxed in the model). Bank fixed terms, CETES and fintechs: not eligible. | `tax_params`, `products_current[exempt_eligible]` |
| Prosofipo protection | 25,000 UDIs per person per SOFIPO, about **220,597 MXN** | `tax_params` |
| IPAB protection (banks) | 400,000 UDIs per person per bank, about 3.53 million MXN | `tax_params` |

Why 200,000 is an interesting amount: it sits just under both the 5-UMA exemption (213,973) and the Prosofipo limit (220,597). One SOFIPO can hold all of it, fully protected and with no withholding. Go to 300,000 and both limits start to bite, so splitting becomes necessary. The report should make that visible.

Simplifications (say these out loud if asked): simple interest, rates held constant except for the optional rate-cut shock, withholding treated as the cost (the real annual tax is computed on real interest and the withholding is credited), fees and subscription costs not included (Meli+, GBM commissions).

---

## 1. Get the data into the Service

1. Run `python3 build_data.py` once. It writes the CSVs in `data/` and `sofipo_savings.xlsx`. The workbook has one sheet per table, each formatted as an Excel table. That matters: the Service turns each Excel table into a model table.
2. Upload `sofipo_savings.xlsx` to your Tec OneDrive (keeping it in OneDrive lets the Service refresh from it later).
3. Go to app.powerbi.com, sign in with your Tec account, open **My workspace**.
4. Choose **New item** (or **+ New**) and pick **Semantic model**, then **Excel**, and select the workbook from OneDrive. The menu names move around; if you cannot find it, search the help for "upload Excel workbook to Power BI service".
5. When it asks, import the data (not "upload to Excel Online"). You should end up with a semantic model called `sofipo_savings` containing all fourteen tables.

Why not upload the CSVs one by one: in the Service each CSV becomes its own separate semantic model, and you cannot relate tables across models without Desktop. The single workbook avoids that.

**If the upload fails or your license blocks it:** check whether your Tec account has a Power BI Pro license (Settings > Manage personal storage shows your license). A free license can still build reports in My workspace, it just cannot share them.

---

## 2. Open the model editor and set relationships

1. In the workspace, open the semantic model's menu and choose **Open data model** (web modeling).
   - If it is greyed out: open **Workspace settings** and enable **Users can edit data models in the Power BI service**. This setting has been in preview; if your tenant blocks it, skip to the note at the end of this section.
2. Check data types: `rate_annual`, `rate_on_excess`, `effective_net_rate`, `target_rate` as **Decimal**, formatted as **Percentage** with 2 decimals. Dates as **Date**.
3. Create these relationships (drag one column onto the other in the model view):
   - `institutions[institution]` 1 to many `products_current[institution]`
   - `institutions[institution]` 1 to many `rates_history[institution]`
   - `institutions[institution]` 1 to many `scenario_allocations[institution]`
4. Leave `amounts`, `horizons`, `tax_params`, `banxico_rate`, `banxico_monthly`, `scenario_summary`, `equity_params`, `equity_scenarios`, `portfolio_mix` and `sp500_backtest` **disconnected**. They feed slicers and lookups, not filters.

If web modeling is not available to you: the precomputed tables (`scenario_summary`, `scenario_allocations`) already contain every result, so sections 4B and 5 still work with plain visuals and slicers. You lose only the live DAX measures of section 3.

---

## 3. DAX measures

In the model editor choose **New measure** and create these on the `products_current` table. Read each one before you paste it; being able to explain them is the point.

```DAX
Selected Amount = SELECTEDVALUE ( amounts[amount_mxn], 200000 )

Selected Days = SELECTEDVALUE ( horizons[horizon_days], 365 )

Withholding Rate =
    LOOKUPVALUE ( tax_params[value], tax_params[param], "isr_withholding_rate" )

SOFIPO Exempt Limit =
    LOOKUPVALUE ( tax_params[value], tax_params[param], "sofipo_exempt_limit_mxn" )
```

Gross interest if you put the whole selected amount in one product. The promo rate only applies up to the product's cap; the rest earns `rate_on_excess` (blank means we could not verify it, so it counts as 0):

```DAX
Gross Interest =
VAR A      = [Selected Amount]
VAR d      = [Selected Days]
VAR r      = SELECTEDVALUE ( products_current[rate_annual] )
VAR cap    = SELECTEDVALUE ( products_current[balance_cap_mxn] )
VAR rx     = COALESCE ( SELECTEDVALUE ( products_current[rate_on_excess] ), 0 )
VAR promo  = IF ( ISBLANK ( cap ), A, MIN ( A, cap ) )
VAR excess = A - promo
RETURN ( promo * r + excess * rx ) * d / 365
```

Tax. Eligible products (see `exempt_eligible`) are withheld only above the exemption; everything else on the whole balance. Treating `unknown` as taxed is the conservative choice; change `"yes"` to `{ "yes", "unknown" }` with `IN` to see the optimistic case:

```DAX
Taxable Balance =
VAR A = [Selected Amount]
RETURN
    IF (
        SELECTEDVALUE ( products_current[exempt_eligible] ) = "yes",
        MAX ( 0, A - [SOFIPO Exempt Limit] ),
        A
    )

ISR Withheld = [Taxable Balance] * [Withholding Rate] * [Selected Days] / 365

Exempt Portion = [Selected Amount] - [Taxable Balance]

Net Return = [Gross Interest] - [ISR Withheld]

Effective Net Rate =
    DIVIDE ( [Net Return], [Selected Amount] ) * 365 / [Selected Days]
```

Protection check (end balance must stay within the limit; blank limit means none reported, or government debt):

```DAX
Protected =
VAR lim = SELECTEDVALUE ( products_current[protection_limit_mxn] )
VAR t   = SELECTEDVALUE ( products_current[type] )
RETURN
    SWITCH (
        TRUE (),
        t = "Government", "Government debt",
        ISBLANK ( lim ), "Not reported",
        [Selected Amount] + [Gross Interest] <= lim, "Yes",
        "No, split it"
    )
```

Only show products that exist for the chosen horizon (on demand, or a fixed term that matches):

```DAX
Available For Horizon =
VAR t = SELECTEDVALUE ( products_current[term_days] )
RETURN IF ( t = 0 || t = [Selected Days], 1, 0 )
```

Best single product for the selection:

```DAX
Best Option =
VAR candidates =
    FILTER (
        ADDCOLUMNS (
            SUMMARIZE ( products_current, products_current[institution], products_current[product] ),
            "@net", [Net Return],
            "@avail", [Available For Horizon]
        ),
        [@avail] = 1
    )
VAR top1 = TOPN ( 1, candidates, [@net], DESC )
RETURN
    CONCATENATEX ( top1, products_current[institution] & " - " & products_current[product], ", " )
```

**Check your measures against Python.** Pick 200,000 and 365 days. The Nu Cajita row must show Gross 13,000.00, ISR 0.00, Net 13,000.00, Effective 6.50% (exempt, under the limit). The CETES 364d row must show Gross 14,480.00, ISR 1,800.00, Net 12,680.00, Effective 6.34%. Products with a monthly fee (Plata+) also need the fee: add `- SELECTEDVALUE ( products_current[monthly_fee_mxn] ) * d / 30.4167` to Gross Interest. The same numbers are in `scenario_summary` (strategies B and A). If yours differ, the usual cause is a rate column stored as text or as a whole number.

---

## 4. Report pages

Create a new report from the semantic model (**Create report**). Build three pages.

### Page 1: "How rates moved"

- **Line chart:** X axis `banxico_monthly[month_end]`, Y axis `banxico_monthly[target_rate]`. Title: "Banxico target rate, 2024 to 2026".
- **Line or scatter chart:** X axis `rates_history[date]`, Y axis average of `rates_history[rate_annual]`, legend `rates_history[institution]`. Add a slicer on `rates_history[term_days]` and set it to 0 (on demand) so you compare like with like. The history is sparse (snapshots in Feb, Jul, Aug and Sep 2026), so turn on markers.
- **Card:** the latest Banxico rate. **Text box:** 2 lines on what you see. Example: Banxico cut from 11.25% to 6.50% since early 2024; fintech base rates followed it down (Nu's on-demand rate went from 7.00% in February to 6.50% by July), while promotional rates on small balances stayed at 13 to 15%.

### Page 2: "Where to put 200,000"

- **Slicers:** `amounts[amount_mxn]` (single select, default 200,000) and `horizons[label]`.
- **Table:** rows `products_current[institution]`, `products_current[product]`, `type`, `rate_annual`, `balance_cap_mxn`, then the measures Gross Interest, ISR Withheld, Net Return, Effective Net Rate, Protected. Add a visual-level filter: Available For Horizon is 1. Sort by Net Return, descending.
- **Card:** Best Option.
- **Conditional formatting:** Protected column red when "No, split it"; `needs_condition` shown so no one confuses a 15% promo with a no-strings rate.

Things to notice: at 200,000 the on-demand promo products look weak, because the 15% only applies to the first 10,000 to 30,000. And banks lose 0.90 points a year to withholding that a SOFIPO under the exemption does not pay.

### Page 3: "Split strategies and rate cuts"

This page uses the precomputed grid, so it works even without web modeling.

- **Slicers:** `scenario_summary[amount_mxn]`, `scenario_summary[horizon_days]`, `scenario_summary[rate_shock_bp]` (0 or 50).
- **Clustered bar chart:** Y axis `strategy`, X axis `net_return_mxn`.
- **Table:** `strategy`, `gross_interest_mxn`, `isr_withheld_mxn`, `net_return_mxn`, `effective_net_rate`, `fully_protected`, `needs_conditions`.
- **Second table or stacked bar** from `scenario_allocations` filtered the same way (sync the slicers): how each strategy splits the money by institution.

The strategies:

| Code | Strategy |
|---|---|
| A | Everything in Stori on demand (SOFIPO, no conditions) |
| B | Everything in Nu Cajita (now a bank, so withholding applies) |
| C | Everything in CETES of the matching term (what you would buy in GBM, before GBM fees) |
| D | Everything in the best SOFIPO fixed term for that horizon (Kubo excluded, see institutions notes) |
| E | Split: fill the capped promo tiers first (Revolut, Ualá, DiDi, Klar Inversión Max, Nu Turbo, Openbank), then the best remaining product after tax and fees (Mifel at 200,000), never above any institution's protection limit |
| F | Same as E but also using Mercado Pago's 15% with Meli+ (protection not reported, subscription cost not included) |

`rate_shock_bp = 50` means on-demand rates drop 0.50 points after day 90, as if Banxico cut twice. Fixed terms keep their rate. Compare A and D at 180 and 365 days with the shock on: that is the value of locking a rate when cuts are expected.

Then set the amount to 300,000 and look at `fully_protected`: A and D become "no", because one SOFIPO would hold more than 220,597. E stays protected by spreading the money across institutions.

### Page 4: "Savings vs the S&P 500"

The recruiter-facing page. It answers: should part of the 200,000 go into the S&P 500 instead, and does it matter whether you hold it with dollar exposure or peso-hedged?

Two ways to own the S&P 500 through GBM:
- **SIC (VOO/IVV):** listed in the Sistema Internacional de Cotizaciones. You pay in pesos, but the value follows the dollar, so a stronger peso eats your return.
- **IVVPESO:** BlackRock's peso-hedged S&P 500 fund on the BMV. No currency risk; the hedge earns roughly the MXN minus USD rate gap (`hedge_carry_assumed` in `equity_params`, an assumption you can edit).

Costs and tax used: GBM commission 0.25% + IVA on the buy and on the sell, fund expense ratio, and 10% ISR on the realized gain (LISR art. 129, paid in the annual return). Losses pay no tax.

Visuals:
- **Table or clustered column, `sp500_backtest`:** period, `spx_total_return_usd`, `usdmxn_change`, `return_in_mxn_sic`, `banxico_avg_target_rate`. The headline: in 2025 the S&P 500 returned 17.9% in dollars, but a SIC investor in pesos made about 2.1%, below the 8.4% average Banxico rate, because the peso gained 13.4%.
- **Matrix, `equity_scenarios`:** rows `spx_scenario` (Bear -15%, Flat, Base +8%, Bull +20%), columns `fx_scenario`, values `net_return`. Slicers `amount_mxn` (200,000) and `vehicle`. Conditional formatting green to red. This is the risk grid.
- **Line or column chart, `portfolio_mix`:** X axis `equity_share` (0%, 10%, 25%, 50%), Y axis `total_net_return`, legend `spx_scenario`, slicers `equity_vehicle` and `fx_scenario`. The savings part uses strategy E on whatever is not in equities.
- **Card with a text box:** the spread between the Bull and Bear outcomes at your chosen equity share. That spread is the price of the extra expected return.

A measure worth writing here, on `portfolio_mix`:

```DAX
Downside vs Savings Only =
VAR worst =
    MINX (
        FILTER ( portfolio_mix, portfolio_mix[spx_scenario] = "Bear" ),
        portfolio_mix[total_net_mxn]
    )
VAR base_only =
    CALCULATE (
        MAX ( portfolio_mix[total_net_mxn] ),
        portfolio_mix[equity_share] = 0,
        REMOVEFILTERS ( portfolio_mix[equity_vehicle], portfolio_mix[fx_scenario], portfolio_mix[spx_scenario] )
    )
RETURN worst - base_only
```

Put `equity_share` on rows and this measure in a table: it shows, in pesos, how much you give up in a bad year for each level of equity exposure.

### What-if parameter (Desktop only) and the workaround

In Desktop you would use Modeling > New parameter to create a numeric slider. That is not available in the web editor. The `amounts` and `horizons` tables are the workaround: they are disconnected tables, and `SELECTEDVALUE` reads the slicer choice with 200,000 and 365 as defaults. To try other amounts, add rows to the `amounts` sheet in the workbook (or to `AMOUNTS` in `build_data.py` for the grid), save it in OneDrive and refresh the semantic model.

---

## 5. Publish and document

1. Save the report. Take one screenshot per page for the GitHub README (`screenshots/` folder).
2. Update the numbers before you show it: rates change often. Rerun the research, edit the lists in `build_data.py`, rebuild, upload.
3. On GitHub, push `README.md`, `GUIDE.md`, `build_data.py`, `data/` and the screenshots. The `.xlsx` is generated, so it is optional.
