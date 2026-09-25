# Where should 200,000 MXN go? Fintech savings, SOFIPOs, CETES and the S&P 500 in Mexico

A Power BI project that compares Mexican fintech, SOFIPO, bank and CETES savings products for a 200,000 MXN balance, and weighs them against the S&P 500 held through GBM, either with dollar exposure (SIC) or peso-hedged (IVVPESO). It looks at how interest rates changed alongside Banxico's cuts, what ISR withholding and the SOFIPO exemption do to net returns, and how to split the money across products and terms while staying inside deposit-protection limits.

Built as a learning project to practice data modeling and DAX in the Power BI Service.

**Status:** the data, model and build guide are done; the Power BI report pages are in progress, and screenshots will be added here as they are finished.

## What it shows

- **How rates moved:** Banxico's target rate from 11.25% (early 2024) to 6.50% (2026), next to snapshots of fintech and SOFIPO rates during 2026.
- **Single-product comparison:** for any amount (default 200,000) and horizon (1, 3, 6, 12 months), the gross interest, ISR withheld, net return, effective net rate and protection status of each product.
- **Split strategies:** six ways to place the money, from "everything in one product" to a split that fills capped promotional tiers first and then the best product after tax and fees. Includes a scenario where on-demand rates fall 0.50 points after 90 days. At 200,000 for a year, the split nets about 12.4% versus 6.3% to 8.7% for any single product.
- **Savings vs the S&P 500:** a 2025 and 2026 backtest in pesos, a bear/flat/base/bull by peso-strength scenario grid, and portfolios with 0% to 50% in equities. In 2025 the S&P 500 returned 17.9% in dollars but about 2.1% for a peso investor through the SIC, because the peso gained 13.4%.

Institutions covered: Nu, Revolut, Ualá, Banco Plata, Openbank, Mifel (banks); Klar, DiDi, Stori, Finsus, Kubo (SOFIPOs); Mercado Pago (fintech); CETES and the S&P 500 via GBM.

## Rules modeled (2026)

| Rule | Value |
|---|---|
| ISR provisional withholding | 0.90% per year on capital (was 0.50% in 2025) |
| 5-UMA interest exemption (LISR art. 93 fr. XX) | Average daily balance up to 5 annual UMAs = 213,973.20 MXN. Covers SOFIPO products and bank on-demand savings accounts where the bank applies it (confirmed for Nu); not bank fixed terms or CETES |
| Equity gains (BMV and SIC) | 10% ISR on the realized gain, paid in the annual return (LISR art. 129) |
| UMA 2026 | 117.31 MXN daily, 42,794.64 MXN annual |
| Prosofipo protection | 25,000 UDIs per person per SOFIPO (about 220,597 MXN on 2026-09-24) |
| IPAB protection | 400,000 UDIs per person per bank (about 3.53 million MXN) |

Notable change: Nu started operating as a bank on 2026-08-06. Its deposits moved from Prosofipo to IPAB protection. Its fixed-term Cajitas are now withheld; the on-demand Cajitas still fall under the savings exemption.

## Repository contents

| Path | What it is |
|---|---|
| `build_data.py` | All sourced figures in one place; builds the CSVs, the scenario grid and the Excel workbook |
| `data/institutions.csv` | Legal type, protection scheme and notes per institution |
| `data/rates_history.csv` | Dated rate observations with caps, conditions and source URLs |
| `data/products_current.csv` | Products available on 2026-09-24, used by the model |
| `data/banxico_rate.csv`, `data/banxico_monthly.csv` | Banxico decisions and a month-end series |
| `data/tax_params.csv` | Tax and protection parameters with sources |
| `data/amounts.csv`, `data/horizons.csv` | Disconnected tables that drive the slicers |
| `data/scenario_summary.csv`, `data/scenario_allocations.csv` | Precomputed results for every amount, horizon, strategy and rate scenario |
| `data/equity_params.csv` | GBM costs, fund fees, hedge carry assumption, S&P 500 returns, USD/MXN levels |
| `data/sp500_backtest.csv` | S&P 500 in dollars vs in pesos, 2025 and 2026 to date, next to the average Banxico rate |
| `data/equity_scenarios.csv` | One-year S&P 500 outcomes by market scenario, peso scenario and vehicle |
| `data/portfolio_mix.csv` | Savings split plus 0% to 50% in the S&P 500, net of costs and taxes |
| `sofipo_savings.xlsx` | Same tables as Excel tables, for upload to the Power BI Service |
| `GUIDE.md` | Step-by-step build instructions, DAX measures and visuals |

## How to reproduce

1. `pip install openpyxl` then `python3 build_data.py`.
2. Upload `sofipo_savings.xlsx` to OneDrive and create a semantic model from it in the Power BI Service.
3. Follow `GUIDE.md` for relationships, measures and report pages.

To update the data, edit the lists at the top of `build_data.py` and rerun it.

## Data sources

All figures were collected on 2026-09-24. Every row in the CSVs has its own `source_url` and a `verified` column.

- Banco de México, monetary policy announcements (target rate history)
- INEGI, UMA 2026 press release
- Secondary tax sources on the 2026 withholding rate and the SOFIPO exemption (Siempre Contable, Russell Bedford, Yahoo Noticias / El Financiero); LISR art. 93 via Justia
- tasas.mx institution pages (Nu, Klar, Revolut, Mercado Pago, market table) and deceroalinfinito.com
- El Financiero (2026-02-14), Expansión (2026-07-24), N+ (Nu's bank conversion), rendimientosmexico.com (Ualá, Banco Plata), tasas.mx (Openbank), El Cronista (Openbank's earlier cap)
- YouTube, Germán Mi Amigo Dinero, monthly rate videos from October 2025 to September 2026 (12 videos, links in `rates_history.csv`). Figures come from the spoken auto-captions only; the on-screen tables are not captured, and ambiguous numbers were skipped. Used to fill the monthly history, resolve Klar's two products, confirm Openbank's cap change and Mercado Pago's above-cap rate, and add Mifel.
- GBM FAQs (commissions, SIC taxation), BlackRock (IVVPESO), First Trust (S&P 500 2025 total return), ChartRow (2026 year to date), El Financiero and DOF (USD/MXN FIX), Trading Economics (current USD/MXN)

**Known gaps:**
- DiDi: the rate on balances above its 10,000 cap was not found in any source, including the videos. The model treats it as 0.
- Mercado Pago: balances above the 25,000 cap earn nothing (Sep 2026 video). Deposit protection is not reported, and its tax treatment is assumed to be the same as a bank's.
- Mifel: the only source is the video series (about 10% up to 500,000). Check the official site before relying on it.
- Bank savings exemption: confirmed only for Nu. Revolut, Ualá, Openbank, Plata and Mifel on-demand accounts are marked `unknown` and taxed, which may understate them.
- S&P 500: the hedge carry (2.5%) is an assumption, the current USD/MXN (17.20) is approximate, US dividend withholding is ignored, and the SIC tax base is assumed to include the currency effect.
- GBM: commissions are not included in the CETES line.
- Kubo Financiero: shown in the data but excluded from the default strategies after users reported blocked withdrawals in May and June 2026.
- The rate history is a series of snapshots, not a continuous daily series.

## Limitations

- Simple interest, and rates held constant except for the optional rate-cut scenario.
- The equity scenarios are illustrations, not forecasts, and carry no probabilities.
- Promotional rates usually require conditions (card spending, subscriptions, account levels). These are flagged but not costed.
- Withholding is treated as the cost. The actual annual tax is computed on real interest (nominal minus inflation), and the withholding is credited against it.

## Disclaimer

This is a personal learning project, not financial, tax or investment advice. Rates, rules and institutions' legal status change often, and some figures here come from secondary sources. Check each institution's official terms, and the SAT and CONDUSEF, before making any decision with real money.
