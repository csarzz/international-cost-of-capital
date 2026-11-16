# International Cost of Capital App

A Streamlit app that compares a project's cost of equity, cost of debt and WACC in US dollars and in the currency of a selected country.

## Formulas

An asterisk marks the foreign currency. Rf and Rf* are the 10-year government bond yields of the United States and of the selected country.

| Step | Formula |
|---|---|
| Debt-to-equity ratio | `D/E = d / (1 − d)`, where d is the target debt share |
| Relevered equity beta | `βE = βU × [1 + (1 − τ) × D/E]` |
| Cost of equity, USD | `Re = Rf + βE × ERP` |
| Pretax cost of debt, USD | `Rd = Rf + s`, where s is the USD BBB spread |
| WACC, USD | `WACC = (1 − d) × Re + d × Rd × (1 − τ)` |
| Conversion factor | `F = (1 + Rf*) / (1 + Rf)` |
| Cost of equity, foreign currency | `Re* = (1 + Re) × F − 1` |
| Foreign spread (automatic) | `s* = s × F` |
| Pretax cost of debt, foreign currency | `Rd* = Rf* + s*` |
| WACC, foreign currency | `WACC* = (1 − d) × Re* + d × Rd* × (1 − τ)` |

The same tax rate τ, debt share d and equity risk premium are used in both currencies.

## What the app does

- Select an industry and a target country, and adjust the equity risk premium and the target debt share.
- Switch the project tax rate between the U.S. rate and the selected country's rate.
- Override any downloaded input: unlevered beta, both corporate tax rates, the USD BBB spread and both 10-year yields.
- Override the foreign-currency BBB spread. When that switch is off, the app shows the automatic value `s × F`.
- See the results side by side, the inputs with their sources and dates, and a step-by-step calculation with the numbers filled in.
- Compare all industries in a chart and a table, and download the project results and the all-industry comparison as CSV files.
- Press **Refresh data** to download all sources again.

Defaults on opening: Software (System & Application), United States as the selected country, the U.S. tax rate, a 6.5% equity risk premium and 30% debt financing.

## Data sources

| Input | Source | How the app reads it |
|---|---|---|
| Unlevered beta by industry | Aswath Damodaran, [betas.xls](https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls) | "Unlevered beta" column of the "Industry Averages" sheet (not the cash-corrected column) |
| USD BBB spread | Aswath Damodaran, [ratings and spreads](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html) | Baa2/BBB row of the table for large non-financial service firms |
| Corporate tax rates | [OECD Corporate Tax Statistics](https://www.oecd.org/en/publications/corporate-tax-statistics-2024_9c27d6e8-en/full-report/statutory-corporate-income-tax-rates_440f9286.html) | Combined (central plus sub-central) statutory rate, latest year, from the OECD data API `sdmx.oecd.org`, dataset `DSD_TAX_CIT@DF_CIT` |
| 10-year government yields | [Trading Economics](https://tradingeconomics.com/bonds) | The "10Y" row of each country's page, `tradingeconomics.com/<country>/government-bond-yield` |

The sources table in the app shows the value used, the date of the data and whether it is live, a snapshot or a manual entry. Betas, the spread and tax rates are cached for 12 hours and yields for 1 hour.

## Notes and limitations

- **OECD.** The OECD publication page is behind a browser check that blocks automated requests, so the app reads the same statutory rates from the OECD's public data API.
- **Trading Economics.** The table on `tradingeconomics.com/bonds` is loaded by a script behind a bot challenge, so the app reads the per-country pages instead, one plain request per country. If Trading Economics declines a request, the app does not try to get around it. It falls back to a dated snapshot for 16 countries, or asks for a manual yield.
- **Snapshots.** If a download fails, the app warns and uses built-in values: Damodaran betas and spread (January 2026), OECD rates (2025 or 2026) and Trading Economics yields for 16 countries (6 November 2026).
- **Countries.** The list holds 64 countries that have an OECD statutory tax rate and are expected to have a 10-year yield page on Trading Economics. Only the United States and Germany pages were tested live.
- **Shared currencies.** Euro-area countries have different government yields, so they give different euro results. The yield gap reflects sovereign risk as well as expected inflation.

## Main assumptions

- Risk is measured against the U.S. equity market, with no separate country risk premium.
- U.S. industry unlevered betas stand in for project risk in any country.
- Relevering assumes debt and the interest tax shield have zero beta.
- Government yields are used as risk-free rates, although they include term and liquidity premia and sometimes sovereign default risk.
- The currency conversion assumes equal real interest rates and relative purchasing power parity, so the yield gap equals expected depreciation.
- The debt share is a constant target and interest is fully tax-deductible when paid.

## How to check the results

Worked example: unlevered beta 0.80, debt share 30%, tax rate 35%, equity risk premium 6.5%, U.S. yield 4%, foreign yield 12%, USD spread 1.5%.

| Rate | USD | Foreign currency |
|---|---|---|
| Cost of equity | 10.6486% | 19.1600% |
| Pretax cost of debt | 5.5000% | 13.6154% |
| After-tax cost of debt | 3.5750% | 8.8500% |
| WACC | 8.5265% | 16.0670% |

Enter these values with the manual overrides to reproduce the table. Three further checks:

1. Equal government yields give equal results in both currencies when the foreign spread override is off.
2. A debt share of zero makes WACC equal to the cost of equity.
3. Overriding the foreign-currency spread leaves the USD results unchanged.

## Files

| File | Purpose |
|---|---|
| `app.py` | The Streamlit app |
| `requirements.txt` | Python packages (tested with Python 3.12) |
| `README.md` | This file |

## Publish on Streamlit Community Cloud

1. Put `app.py`, `requirements.txt` and `README.md` at the top level of a public GitHub repository.
2. At [share.streamlit.io](https://share.streamlit.io) choose **Create app**, then select the repository, branch `main` and file `app.py`.
3. Under **Advanced settings** choose Python 3.12 and leave Secrets empty, then click **Deploy**.

## Run locally (optional)

```
pip install -r requirements.txt
streamlit run app.py
```

## Credits

Code written by Claude (Anthropic, model Claude Fable 5.2) on 16 November 2026 for a Kellogg International Finance assignment. Formulas follow the course material. Data: Aswath Damodaran, OECD, Trading Economics.
