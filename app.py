"""International cost of capital app.

Compares a project's cost of equity, cost of debt and WACC in US dollars and
in the currency of a selected country.

Formulas (an asterisk marks the foreign currency)
-------------------------------------------------
D/E        = d / (1 - d)                         d = target debt share
beta_E     = beta_U x [1 + (1 - tax) x D/E]      relevered equity beta
Re         = Rf + beta_E x ERP                   cost of equity in USD (CAPM)
Rd         = Rf + s                              pretax cost of debt in USD
WACC       = (1 - d) x Re + d x Rd x (1 - tax)
F          = (1 + Rf*) / (1 + Rf)                currency conversion factor
Re*        = (1 + Re) x F - 1                    cost of equity in foreign currency
s*         = s x F                               foreign spread (automatic)
Rd*        = Rf* + s*
WACC*      = (1 - d) x Re* + d x Rd* x (1 - tax)

Rf and Rf* are the 10-year government bond yields of the United States and of
the selected country. The same tax rate, debt share and equity risk premium
are used in both currencies.

Compatible with Python 3.12.
"""

from __future__ import annotations

import csv
import html
import io
import re
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import xlrd

# --------------------------------------------------------------------------
# Sources
# --------------------------------------------------------------------------
BETAS_URL = "https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls"
RATINGS_URL = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html"
OECD_PAGE_URL = (
    "https://www.oecd.org/en/publications/corporate-tax-statistics-2024_9c27d6e8-en/"
    "full-report/statutory-corporate-income-tax-rates_440f9286.html"
)
# The OECD publication page sits behind a browser check, so the app reads the
# same statutory rates from the OECD's public data API (Data Explorer).
OECD_API_URL = (
    "https://sdmx.oecd.org/public/rest/data/OECD.CTP.TPS,DSD_TAX_CIT@DF_CIT,/"
    ".A.CIT_C.ST..S13...?startPeriod=2020"
)
TE_BONDS_URL = "https://tradingeconomics.com/bonds"
TE_COUNTRY_URL = "https://tradingeconomics.com/{slug}/government-bond-yield"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
TIMEOUT = 20                       # seconds per request
SLOW_CACHE = 12 * 60 * 60          # betas, spread, tax rates: 12 hours
YIELD_CACHE = 60 * 60              # bond yields: 1 hour

DEFAULT_INDUSTRY = "Software (System & Application)"
DEFAULT_COUNTRY = "USA"
DEFAULT_ERP = 6.5                  # percent
DEFAULT_DEBT_SHARE = 30            # percent
US = "USA"

# Countries offered in the app: those with a statutory corporate tax rate in
# the OECD data and a 10-year government bond page on Trading Economics.
#   ISO3 code: (name, Trading Economics page name, currency code)
COUNTRIES = {
    "AUS": ("Australia", "australia", "AUD"),
    "AUT": ("Austria", "austria", "EUR"),
    "BEL": ("Belgium", "belgium", "EUR"),
    "BWA": ("Botswana", "botswana", "BWP"),
    "BRA": ("Brazil", "brazil", "BRL"),
    "BGR": ("Bulgaria", "bulgaria", "EUR"),
    "CAN": ("Canada", "canada", "CAD"),
    "CHL": ("Chile", "chile", "CLP"),
    "CHN": ("China", "china", "CNY"),
    "COL": ("Colombia", "colombia", "COP"),
    "HRV": ("Croatia", "croatia", "EUR"),
    "CZE": ("Czech Republic", "czech-republic", "CZK"),
    "DNK": ("Denmark", "denmark", "DKK"),
    "EGY": ("Egypt", "egypt", "EGP"),
    "FIN": ("Finland", "finland", "EUR"),
    "FRA": ("France", "france", "EUR"),
    "DEU": ("Germany", "germany", "EUR"),
    "GRC": ("Greece", "greece", "EUR"),
    "HKG": ("Hong Kong", "hong-kong", "HKD"),
    "HUN": ("Hungary", "hungary", "HUF"),
    "ISL": ("Iceland", "iceland", "ISK"),
    "IND": ("India", "india", "INR"),
    "IDN": ("Indonesia", "indonesia", "IDR"),
    "IRL": ("Ireland", "ireland", "EUR"),
    "ISR": ("Israel", "israel", "ILS"),
    "ITA": ("Italy", "italy", "EUR"),
    "JPN": ("Japan", "japan", "JPY"),
    "KAZ": ("Kazakhstan", "kazakhstan", "KZT"),
    "KEN": ("Kenya", "kenya", "KES"),
    "LTU": ("Lithuania", "lithuania", "EUR"),
    "MYS": ("Malaysia", "malaysia", "MYR"),
    "MLT": ("Malta", "malta", "EUR"),
    "MUS": ("Mauritius", "mauritius", "MUR"),
    "MEX": ("Mexico", "mexico", "MXN"),
    "MAR": ("Morocco", "morocco", "MAD"),
    "NAM": ("Namibia", "namibia", "NAD"),
    "NLD": ("Netherlands", "netherlands", "EUR"),
    "NZL": ("New Zealand", "new-zealand", "NZD"),
    "NGA": ("Nigeria", "nigeria", "NGN"),
    "NOR": ("Norway", "norway", "NOK"),
    "PAK": ("Pakistan", "pakistan", "PKR"),
    "PER": ("Peru", "peru", "PEN"),
    "PHL": ("Philippines", "philippines", "PHP"),
    "POL": ("Poland", "poland", "PLN"),
    "PRT": ("Portugal", "portugal", "EUR"),
    "QAT": ("Qatar", "qatar", "QAR"),
    "ROU": ("Romania", "romania", "RON"),
    "SRB": ("Serbia", "serbia", "RSD"),
    "SGP": ("Singapore", "singapore", "SGD"),
    "SVK": ("Slovakia", "slovakia", "EUR"),
    "SVN": ("Slovenia", "slovenia", "EUR"),
    "ZAF": ("South Africa", "south-africa", "ZAR"),
    "KOR": ("South Korea", "south-korea", "KRW"),
    "ESP": ("Spain", "spain", "EUR"),
    "LKA": ("Sri Lanka", "sri-lanka", "LKR"),
    "SWE": ("Sweden", "sweden", "SEK"),
    "CHE": ("Switzerland", "switzerland", "CHF"),
    "THA": ("Thailand", "thailand", "THB"),
    "TUR": ("Turkey", "turkey", "TRY"),
    "UKR": ("Ukraine", "ukraine", "UAH"),
    "GBR": ("United Kingdom", "united-kingdom", "GBP"),
    "USA": ("United States", "united-states", "USD"),
    "VNM": ("Vietnam", "vietnam", "VND"),
    "ZMB": ("Zambia", "zambia", "ZMW"),
}

# --------------------------------------------------------------------------
# Snapshots used only when a live download fails (shown with their dates)
# --------------------------------------------------------------------------
BETAS_SNAPSHOT_DATE = "5 January 2026"
# industry: unlevered beta
BETAS_SNAPSHOT = {
    "Advertising": 0.930086,
    "Aerospace/Defense": 0.846669,
    "Air Transport": 0.704050,
    "Apparel": 0.757982,
    "Auto & Truck": 1.269033,
    "Auto Parts": 1.021714,
    "Bank (Money Center)": 0.341059,
    "Banks (Regional)": 0.286516,
    "Beverage (Alcoholic)": 0.613172,
    "Beverage (Soft)": 0.555671,
    "Broadcasting": 0.286362,
    "Brokerage & Investment Banking": 0.580849,
    "Building Materials": 0.930162,
    "Business & Consumer Services": 0.773588,
    "Cable TV": 0.352392,
    "Chemical (Basic)": 0.580033,
    "Chemical (Diversified)": 0.366552,
    "Chemical (Specialty)": 0.792227,
    "Coal & Related Energy": 1.016535,
    "Computer Services": 0.915510,
    "Computers/Peripherals": 1.305062,
    "Construction Supplies": 1.016068,
    "Diversified": 0.788930,
    "Drugs (Biotechnology)": 1.034129,
    "Drugs (Pharmaceutical)": 0.886211,
    "Education": 0.660196,
    "Electrical Equipment": 1.147795,
    "Electronics (Consumer & Office)": 0.829853,
    "Electronics (General)": 0.897365,
    "Engineering/Construction": 1.094912,
    "Entertainment": 0.737184,
    "Environmental & Waste Services": 0.814469,
    "Farming/Agriculture": 0.813020,
    "Financial Svcs. (Non-bank & Insurance)": 0.318882,
    "Food Processing": 0.457418,
    "Food Wholesalers": 0.640882,
    "Furn/Home Furnishings": 0.624788,
    "Green & Renewable Energy": 0.463301,
    "Healthcare Products": 0.829063,
    "Healthcare Support Services": 0.688916,
    "Heathcare Information and Technology": 0.991301,
    "Homebuilding": 0.784960,
    "Hospitals/Healthcare Facilities": 0.551730,
    "Hotel/Gaming": 0.832251,
    "Household Products": 0.717816,
    "Information Services": 0.737177,
    "Insurance (General)": 0.563640,
    "Insurance (Life)": 0.427156,
    "Insurance (Prop/Cas.)": 0.435656,
    "Investments & Asset Management": 0.529716,
    "Machinery": 0.867976,
    "Metals & Mining": 0.963643,
    "Office Equipment & Services": 0.980604,
    "Oil/Gas (Integrated)": 0.271219,
    "Oil/Gas (Production and Exploration)": 0.562760,
    "Oil/Gas Distribution": 0.465076,
    "Oilfield Svcs/Equip.": 0.743039,
    "Packaging & Container": 0.723069,
    "Paper/Forest Products": 0.721553,
    "Power": 0.310024,
    "Precious Metals": 0.793243,
    "Publishing & Newspapers": 0.477307,
    "R.E.I.T.": 0.392904,
    "Real Estate (Development)": 0.478288,
    "Real Estate (General/Diversified)": 0.577859,
    "Real Estate (Operations & Services)": 0.814892,
    "Recreation": 0.695091,
    "Reinsurance": 0.438409,
    "Restaurant/Dining": 0.767420,
    "Retail (Automotive)": 0.698433,
    "Retail (Building Supply)": 1.307014,
    "Retail (Distributors)": 0.782683,
    "Retail (General)": 0.759964,
    "Retail (Grocery and Food)": 0.804751,
    "Retail (REITs)": 0.436173,
    "Retail (Special Lines)": 0.948783,
    "Rubber& Tires": 0.143534,
    "Semiconductor": 1.489258,
    "Semiconductor Equip": 1.348073,
    "Shipbuilding & Marine": 0.644013,
    "Shoe": 0.933675,
    "Software (Entertainment)": 1.012781,
    "Software (Internet)": 1.546064,
    "Software (System & Application)": 1.225392,
    "Steel": 0.903514,
    "Telecom (Wireless)": 0.387259,
    "Telecom. Equipment": 0.863491,
    "Telecom. Services": 0.365434,
    "Tobacco": 0.677590,
    "Transportation": 0.675311,
    "Transportation (Railroads)": 0.806948,
    "Trucking": 0.850418,
    "Utility (General)": 0.148613,
    "Utility (Water)": 0.281177,
}

SPREAD_SNAPSHOT = 1.11             # percent, Baa2/BBB, large non-financial firms
SPREAD_SNAPSHOT_DATE = "January 2026"

# ISO3 code: (combined statutory corporate tax rate in percent, year)
TAX_SNAPSHOT = {
    "AUS": (30, 2026),
    "AUT": (23, 2026),
    "BEL": (25, 2026),
    "BWA": (22, 2026),
    "BRA": (34, 2026),
    "BGR": (10, 2026),
    "CAN": (26, 2026),
    "CHL": (27, 2026),
    "CHN": (25, 2026),
    "COL": (35, 2025),
    "HRV": (18, 2026),
    "CZE": (21, 2026),
    "DNK": (22, 2026),
    "EGY": (22.5, 2026),
    "FIN": (20, 2026),
    "FRA": (36.13, 2026),
    "DEU": (30.133, 2026),
    "GRC": (22, 2026),
    "HKG": (16.5, 2026),
    "HUN": (9, 2026),
    "ISL": (20, 2026),
    "IND": (25.17, 2026),
    "IDN": (22, 2026),
    "IRL": (12.5, 2026),
    "ISR": (23, 2025),
    "ITA": (27.81, 2026),
    "JPN": (29.74, 2026),
    "KAZ": (20, 2026),
    "KEN": (30, 2026),
    "LTU": (17, 2026),
    "MYS": (24, 2026),
    "MLT": (35, 2026),
    "MUS": (15, 2026),
    "MEX": (30, 2026),
    "MAR": (35, 2026),
    "NAM": (32, 2026),
    "NLD": (25.8, 2026),
    "NZL": (28, 2026),
    "NGA": (30, 2026),
    "NOR": (22, 2026),
    "PAK": (29, 2026),
    "PER": (29.5, 2026),
    "PHL": (25, 2026),
    "POL": (19, 2026),
    "PRT": (29.5, 2026),
    "QAT": (10, 2026),
    "ROU": (16, 2026),
    "SRB": (15, 2026),
    "SGP": (17, 2026),
    "SVK": (24, 2026),
    "SVN": (22, 2026),
    "ZAF": (27, 2026),
    "KOR": (27.5, 2026),
    "ESP": (25, 2026),
    "LKA": (30, 2026),
    "SWE": (20.6, 2026),
    "CHE": (19.4719, 2026),
    "THA": (20, 2026),
    "TUR": (25, 2025),
    "UKR": (18, 2026),
    "GBR": (25, 2026),
    "USA": (25.5285, 2026),
    "VNM": (20, 2026),
    "ZMB": (30, 2026),
}

# 10-year government yields in percent from Trading Economics, as reported in
# the course packet. ISO3 code: yield
YIELD_SNAPSHOT_DATE = "6 November 2026"
YIELD_SNAPSHOT = {
    "BRA": 15.2, "COL": 12.8, "MEX": 9.1, "IND": 7.2, "POL": 6.2, "GBR": 5.2,
    "AUS": 5.0, "USA": 4.9, "FRA": 4.8, "SVK": 4.2, "PRT": 3.7, "CAN": 3.7,
    "DEU": 3.3, "JPN": 2.6, "CHN": 1.6, "CHE": 0.4,
}

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def _now() -> str:
    return datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC")


def _failure(exc: Exception) -> str:
    text = str(exc).strip()
    return f"{type(exc).__name__}: {text[:160]}" if text else type(exc).__name__


def _cells(row_html: str) -> list[str]:
    """Plain text of each cell in one HTML table row."""
    cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", row_html, flags=re.S | re.I)
    return [re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", c))).strip() for c in cells]


def _rows(page: str) -> list[list[str]]:
    return [_cells(r) for r in re.findall(r"<tr[^>]*>(.*?)</tr>", page, flags=re.S | re.I)]


def pct(value: float, digits: int = 2) -> str:
    return f"{value * 100:.{digits}f}%"


# --------------------------------------------------------------------------
# Parsers (kept separate from the downloads so they can be tested offline)
# --------------------------------------------------------------------------
def parse_betas(content: bytes) -> tuple[dict[str, float], str]:
    """Unlevered beta by industry from Damodaran's betas.xls, plus its date."""
    book = xlrd.open_workbook(file_contents=content)
    sheet = book.sheet_by_name("Industry Averages")
    header_row = name_col = beta_col = None
    updated = ""
    for r in range(min(sheet.nrows, 40)):
        values = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        lowered = [v.lower() for v in values]
        if lowered and lowered[0].startswith("date updated") and sheet.ncols > 1:
            raw = sheet.cell_value(r, 1)
            if isinstance(raw, float):
                stamp = xlrd.xldate_as_datetime(raw, book.datemode)
                updated = f"{stamp.day} {stamp:%B %Y}"
            else:
                updated = str(raw)
        if "industry name" in lowered and "unlevered beta" in lowered:
            header_row = r
            name_col = lowered.index("industry name")
            beta_col = lowered.index("unlevered beta")       # not the cash-corrected column
            break
    if header_row is None:
        raise ValueError("could not find the 'Unlevered beta' column")

    betas: dict[str, float] = {}
    for r in range(header_row + 1, sheet.nrows):
        name = str(sheet.cell_value(r, name_col)).strip()
        value = sheet.cell_value(r, beta_col)
        if not name or name.lower().startswith("total market") or not isinstance(value, float):
            continue
        betas[name] = float(value)
    if len(betas) < 50:
        raise ValueError(f"only {len(betas)} industries found")
    return betas, updated


def parse_bbb_spread(page: str) -> tuple[float, str]:
    """Baa2/BBB spread (percent) for large non-financial firms from ratings.html."""
    spread = None
    for cells in _rows(page):
        for i, cell in enumerate(cells):
            if "Baa2/BBB" in cell:                 # first match in a row is the large-firm block
                for later in cells[i + 1:]:
                    match = re.fullmatch(r"(-?\d+(?:\.\d+)?)\s*%", later)
                    if match:
                        spread = float(match.group(1))
                        break
                break
        if spread is not None:
            break
    if spread is None or not 0 < spread < 20:
        raise ValueError("could not read the Baa2/BBB spread")
    text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", page)))
    dated = re.search(r"as of\s+([A-Z][a-z]+\s+\d{4})", text)
    return spread, dated.group(1) if dated else ""


def parse_tax_rates(text: str) -> dict[str, tuple[float, int]]:
    """Latest combined statutory corporate tax rate per country from the OECD CSV."""
    rates: dict[str, tuple[float, int]] = {}
    for row in csv.DictReader(io.StringIO(text.lstrip("﻿"))):
        try:
            code = row["REF_AREA"].split(":")[0].strip()
            year = int(row["TIME_PERIOD"])
            rate = float(row["OBS_VALUE"])
        except (KeyError, ValueError, TypeError):
            continue
        if code not in rates or year > rates[code][1]:
            rates[code] = (rate, year)
    if len(rates) < 30:
        raise ValueError(f"only {len(rates)} countries found")
    return rates


def parse_yield(page: str, today: date | None = None) -> tuple[float, str]:
    """10-year government bond yield (percent) and its date from a Trading Economics page."""
    today = today or datetime.now(timezone.utc).date()
    for cells in _rows(page):
        cells = [c for c in cells if c]
        if len(cells) < 2 or not re.search(r"\b10Y$", cells[0]):
            continue
        try:
            value = float(cells[1].replace(",", ""))
        except ValueError:
            continue
        if not -5 < value < 150:
            continue
        when = ""
        stamp = re.fullmatch(r"([A-Z][a-z]{2})/(\d{1,2})", cells[-1])
        if stamp and stamp.group(1) in MONTHS:
            month, day = MONTHS[stamp.group(1)], int(stamp.group(2))
            year = today.year
            try:
                if date(year, month, day) > today + timedelta(days=2):
                    year -= 1                       # a December quote read in January
                when = f"{day} {stamp.group(1)} {year}"
            except ValueError:
                when = cells[-1]
        return value, when
    raise ValueError("no 10-year yield found on the page")


# --------------------------------------------------------------------------
# Downloads. Each returns a dict and never raises, so a failed source falls
# back to its snapshot. Results are cached; "Refresh data" clears the cache.
# --------------------------------------------------------------------------
@st.cache_data(ttl=SLOW_CACHE, show_spinner=False)
def load_betas() -> dict:
    try:
        response = requests.get(BETAS_URL, headers=HEADERS, timeout=TIMEOUT)
        response.raise_for_status()
        betas, updated = parse_betas(response.content)
        return {"ok": True, "betas": betas, "date": updated, "fetched": _now(), "error": ""}
    except Exception as exc:  # noqa: BLE001 - any failure means "use the snapshot"
        return {"ok": False, "betas": dict(BETAS_SNAPSHOT), "date": BETAS_SNAPSHOT_DATE,
                "fetched": _now(), "error": _failure(exc)}


@st.cache_data(ttl=SLOW_CACHE, show_spinner=False)
def load_bbb_spread() -> dict:
    try:
        response = requests.get(RATINGS_URL, headers=HEADERS, timeout=TIMEOUT)
        response.raise_for_status()
        spread, dated = parse_bbb_spread(response.content.decode("latin-1"))
        return {"ok": True, "spread": spread, "date": dated, "fetched": _now(), "error": ""}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "spread": SPREAD_SNAPSHOT, "date": SPREAD_SNAPSHOT_DATE,
                "fetched": _now(), "error": _failure(exc)}


@st.cache_data(ttl=SLOW_CACHE, show_spinner=False)
def load_tax_rates() -> dict:
    try:
        response = requests.get(
            OECD_API_URL,
            headers={"User-Agent": HEADERS["User-Agent"],
                     "Accept": "application/vnd.sdmx.data+csv; charset=utf-8"},
            timeout=TIMEOUT,
        )
        response.raise_for_status()
        rates = parse_tax_rates(response.content.decode("utf-8"))
        return {"ok": True, "rates": rates, "fetched": _now(), "error": ""}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "rates": dict(TAX_SNAPSHOT), "fetched": _now(), "error": _failure(exc)}


@st.cache_data(ttl=YIELD_CACHE, show_spinner=False)
def load_yield(code: str) -> dict:
    """10-year yield of one country from its Trading Economics page.

    One plain page request per country. If Trading Economics declines the
    request the app does not try to get around it: it uses the snapshot, or
    asks for a manual entry when there is none.
    """
    try:
        response = requests.get(
            TE_COUNTRY_URL.format(slug=COUNTRIES[code][1]), headers=HEADERS, timeout=TIMEOUT
        )
        if response.status_code != 200 or response.headers.get("x-amzn-waf-action"):
            raise RuntimeError(f"Trading Economics declined the request (HTTP {response.status_code})")
        value, when = parse_yield(response.text)
        return {"ok": True, "yield": value, "date": when, "fetched": _now(), "error": ""}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "yield": YIELD_SNAPSHOT.get(code), "date": YIELD_SNAPSHOT_DATE,
                "fetched": _now(), "error": _failure(exc)}


# --------------------------------------------------------------------------
# Calculations (all rates as decimals, e.g. 0.065 for 6.5%)
# --------------------------------------------------------------------------
def cost_of_capital(beta_u, debt_share, tax, erp, rf_us, rf_foreign, spread_us,
                    spread_foreign_override=None) -> dict:
    """Cost of equity, cost of debt and WACC in USD and in the foreign currency.

    Works on single numbers and on pandas Series of unlevered betas.
    """
    d = debt_share
    debt_to_equity = d / (1 - d)
    beta_e = beta_u * (1 + (1 - tax) * debt_to_equity)

    # US dollars
    re_usd = rf_us + beta_e * erp
    rd_usd = rf_us + spread_us
    wacc_usd = (1 - d) * re_usd + d * rd_usd * (1 - tax)

    # Foreign currency
    factor = (1 + rf_foreign) / (1 + rf_us)
    re_foreign = (1 + re_usd) * factor - 1
    spread_auto = spread_us * factor
    spread_foreign = spread_auto if spread_foreign_override is None else spread_foreign_override
    rd_foreign = rf_foreign + spread_foreign
    wacc_foreign = (1 - d) * re_foreign + d * rd_foreign * (1 - tax)

    return {
        "debt_to_equity": debt_to_equity, "beta_e": beta_e, "factor": factor,
        "re_usd": re_usd, "rd_usd": rd_usd, "rd_usd_after_tax": rd_usd * (1 - tax),
        "wacc_usd": wacc_usd,
        "re_foreign": re_foreign, "spread_auto": spread_auto, "spread_foreign": spread_foreign,
        "rd_foreign": rd_foreign, "rd_foreign_after_tax": rd_foreign * (1 - tax),
        "wacc_foreign": wacc_foreign,
    }


# --------------------------------------------------------------------------
# Chart
# --------------------------------------------------------------------------
def industry_chart(table: pd.DataFrame, column: str, label: str, selected: str) -> go.Figure:
    """Horizontal bars of WACC by industry, with the selected industry highlighted."""
    try:
        dark = st.context.theme.type == "dark"
    except Exception:  # noqa: BLE001
        dark = False
    base, accent = ("#3987e5", "#d95926") if dark else ("#2a78d6", "#eb6834")

    data = table.sort_values(column)
    is_selected = data["Industry"] == selected
    figure = go.Figure(
        go.Bar(
            x=data[column] * 100,
            y=data["Industry"],
            orientation="h",
            width=0.6,
            marker={"color": [accent if hit else base for hit in is_selected],
                    "cornerradius": 4},
            hovertemplate="%{y}<br>" + label + ": %{x:.2f}%<extra></extra>",
        )
    )
    if is_selected.any():
        row = data[is_selected].iloc[0]
        figure.add_annotation(
            x=row[column] * 100, y=row["Industry"], xanchor="left", xshift=6, showarrow=False,
            text=f"Selected: {row[column] * 100:.2f}%",
        )
    figure.update_layout(
        height=max(400, 20 * len(data) + 80),
        margin={"l": 10, "r": 30, "t": 60, "b": 10},
        xaxis={"title": {"text": f"{label} (%)"}, "ticksuffix": "%", "rangemode": "tozero",
               "side": "top"},
        yaxis={"automargin": True, "tickfont": {"size": 11}},
        showlegend=False,
    )
    return figure


# --------------------------------------------------------------------------
# Page
# --------------------------------------------------------------------------
def override(container, label: str, fetched: float | None, key: str, *,
             step: float, fmt: str, low: float, high: float, help_text: str = "") -> float | None:
    """A switch plus a number box. Returns the manual value, or None when the switch is off."""
    if not container.toggle(label, key=f"use_{key}", help=help_text or None):
        return None
    start = fetched if fetched is not None else 0.0
    return container.number_input(
        label, min_value=low, max_value=high, value=float(start), step=step, format=fmt,
        key=f"value_{key}", label_visibility="collapsed",
    )


def main() -> None:
    st.set_page_config(page_title="International Cost of Capital", layout="wide")
    st.title("International cost of capital")
    st.caption(
        "Cost of equity, cost of debt and WACC for a project, in US dollars and in the "
        "currency of a selected country."
    )

    side = st.sidebar
    if side.button("Refresh data", help="Download all four sources again."):
        st.cache_data.clear()
        st.rerun()

    # ---- downloads that do not depend on the country ---------------------
    with st.spinner("Downloading data…"):
        beta_data = load_betas()
        spread_data = load_bbb_spread()
        tax_data = load_tax_rates()

    # ---- project choices --------------------------------------------------
    side.header("Project")
    industries = list(beta_data["betas"])
    industry = side.selectbox(
        "Industry", industries,
        index=industries.index(DEFAULT_INDUSTRY) if DEFAULT_INDUSTRY in industries else 0,
    )
    codes = sorted(COUNTRIES, key=lambda c: COUNTRIES[c][0])
    code = side.selectbox(
        "Target country", codes, index=codes.index(DEFAULT_COUNTRY),
        format_func=lambda c: f"{COUNTRIES[c][0]} ({COUNTRIES[c][2]})",
    )
    country, _, currency = COUNTRIES[code]
    is_us = code == US
    foreign_label = f"{currency} ({country})"

    erp_pct = side.number_input("Equity risk premium (%)", min_value=0.0, max_value=25.0,
                                value=DEFAULT_ERP, step=0.1, format="%.2f")
    debt_pct = side.slider("Target debt share, D/(D+E) (%)", min_value=0, max_value=95,
                           value=DEFAULT_DEBT_SHARE, step=1)

    with st.spinner("Downloading bond yields…"):
        yield_us = load_yield(US)
        yield_foreign = yield_us if is_us else load_yield(code)

    def tax_lookup(c: str) -> tuple[float | None, str, bool]:
        """(rate in percent, year, live?) for one country."""
        if c in tax_data["rates"]:
            rate, year = tax_data["rates"][c]
            return rate, str(year), tax_data["ok"]
        if c in TAX_SNAPSHOT:
            rate, year = TAX_SNAPSHOT[c]
            return rate, str(year), False
        return None, "", False

    tax_us_fetched, tax_us_year, tax_us_live = tax_lookup(US)
    tax_f_fetched, tax_f_year, tax_f_live = tax_lookup(code)

    # ---- manual overrides -------------------------------------------------
    side.header("Manual overrides")
    side.caption("Switch one on to replace the downloaded value.")
    beta_manual = override(side, "Unlevered beta", beta_data["betas"][industry], "beta",
                           step=0.05, fmt="%.4f", low=0.0, high=5.0)
    tax_us_manual = override(side, "U.S. corporate tax rate (%)", tax_us_fetched, "tax_us",
                             step=0.5, fmt="%.2f", low=0.0, high=80.0)
    tax_f_manual = None
    if not is_us:
        tax_f_manual = override(side, f"{country} corporate tax rate (%)", tax_f_fetched,
                                f"tax_{code}", step=0.5, fmt="%.2f", low=0.0, high=80.0)
    spread_manual = override(side, "USD BBB spread (%)", spread_data["spread"], "spread_us",
                             step=0.05, fmt="%.2f", low=0.0, high=30.0)
    rf_us_manual = override(side, "U.S. 10-year yield (%)", yield_us["yield"], "rf_us",
                            step=0.05, fmt="%.2f", low=-5.0, high=150.0)
    rf_f_manual = None
    if not is_us:
        rf_f_manual = override(side, f"{country} 10-year yield (%)", yield_foreign["yield"],
                               f"rf_{code}", step=0.05, fmt="%.2f", low=-5.0, high=150.0)

    # ---- values used -------------------------------------------------------
    def used(manual, fetched):
        return manual if manual is not None else fetched

    beta_u = used(beta_manual, beta_data["betas"][industry])
    tax_us_pct = used(tax_us_manual, tax_us_fetched)
    tax_f_pct = tax_us_pct if is_us else used(tax_f_manual, tax_f_fetched)
    spread_pct = used(spread_manual, spread_data["spread"])
    rf_us_pct = used(rf_us_manual, yield_us["yield"])
    rf_f_pct = rf_us_pct if is_us else used(rf_f_manual, yield_foreign["yield"])

    missing = []
    if rf_us_pct is None:
        missing.append("the U.S. 10-year yield")
    if rf_f_pct is None:
        missing.append(f"the {country} 10-year yield")
    if tax_us_pct is None:
        missing.append("the U.S. corporate tax rate")
    if tax_f_pct is None:
        missing.append(f"the {country} corporate tax rate")
    if missing:
        st.error(
            "No value is available for " + " and ".join(missing) + ". The live download "
            "failed and there is no snapshot for it. Switch on the matching manual override "
            "in the sidebar and enter a value, or press Refresh data."
        )
        st.stop()

    # The automatic foreign spread needs the yields, so this switch comes last.
    spread_auto_pct = spread_pct * (1 + rf_f_pct / 100) / (1 + rf_us_pct / 100)
    spread_f_manual = override(
        side, "Foreign-currency BBB spread (%)", round(spread_auto_pct, 4), f"spread_{code}",
        step=0.05, fmt="%.4f", low=0.0, high=60.0,
        help_text="Off: the app converts the USD spread with the same factor as the cost of equity.",
    )
    if spread_f_manual is None:
        side.caption(
            f"Automatic {currency} spread: {spread_auto_pct:.4f}% = {spread_pct:.2f}% × "
            f"(1 + {rf_f_pct:.2f}%) / (1 + {rf_us_pct:.2f}%)"
        )

    # ---- project tax rate --------------------------------------------------
    if is_us:
        tax_choice = "U.S."
        st.caption(f"Project tax rate: U.S. rate, {tax_us_pct:.2f}%.")
    else:
        options = ["U.S.", country]
        tax_choice = st.radio(
            "Project tax rate (used in both currencies)", options, horizontal=True,
            format_func=lambda o: f"{o} rate: {(tax_us_pct if o == 'U.S.' else tax_f_pct):.2f}%",
        )
    tax_pct = tax_us_pct if tax_choice == "U.S." else tax_f_pct

    # ---- calculation -------------------------------------------------------
    inputs = {
        "debt_share": debt_pct / 100, "tax": tax_pct / 100, "erp": erp_pct / 100,
        "rf_us": rf_us_pct / 100, "rf_foreign": rf_f_pct / 100, "spread_us": spread_pct / 100,
        "spread_foreign_override": None if spread_f_manual is None else spread_f_manual / 100,
    }
    result = cost_of_capital(beta_u, **inputs)

    # ---- warnings about fallbacks -----------------------------------------
    problems = []
    if not beta_data["ok"] and beta_manual is None:
        problems.append(f"Damodaran industry betas ({beta_data['error']})")
    if not spread_data["ok"] and spread_manual is None:
        problems.append(f"Damodaran BBB spread ({spread_data['error']})")
    if not tax_data["ok"] and (tax_us_manual is None or (not is_us and tax_f_manual is None)):
        problems.append(f"OECD tax rates ({tax_data['error']})")
    if not yield_us["ok"] and rf_us_manual is None:
        problems.append(f"U.S. 10-year yield ({yield_us['error']})")
    if not is_us and not yield_foreign["ok"] and rf_f_manual is None:
        problems.append(f"{country} 10-year yield ({yield_foreign['error']})")
    if problems:
        st.warning(
            "Live download failed for: " + "; ".join(problems) + ". The app is using the "
            "dated snapshot shown in the sources table. Press Refresh data to try again, or "
            "enter current values with the manual overrides."
        )

    # ---- results side by side ----------------------------------------------
    st.subheader("Results")
    left, right = st.columns(2)
    for column, title, key in ((left, "USD", "usd"), (right, foreign_label, "foreign")):
        box = column.container(border=True)
        box.markdown(f"**{title}**")
        a, b, c = box.columns(3)
        a.metric("Cost of equity", pct(result[f"re_{key}"]))
        b.metric("After-tax cost of debt", pct(result[f"rd_{key}_after_tax"]))
        c.metric("WACC", pct(result[f"wacc_{key}"]))

    spread_f_note = "manual" if spread_f_manual is not None else "automatic"
    results_table = pd.DataFrame(
        {
            "Item": [
                "10-year government yield (risk-free proxy)",
                "BBB borrowing spread",
                "Relevered equity beta",
                "Cost of equity",
                "Pretax cost of debt",
                "After-tax cost of debt",
                "WACC",
            ],
            "USD": [
                pct(inputs["rf_us"], 4), pct(inputs["spread_us"], 4), f"{result['beta_e']:.4f}",
                pct(result["re_usd"], 4), pct(result["rd_usd"], 4),
                pct(result["rd_usd_after_tax"], 4), pct(result["wacc_usd"], 4),
            ],
            foreign_label: [
                pct(inputs["rf_foreign"], 4),
                f"{pct(result['spread_foreign'], 4)} ({spread_f_note})",
                f"{result['beta_e']:.4f}",
                pct(result["re_foreign"], 4), pct(result["rd_foreign"], 4),
                pct(result["rd_foreign_after_tax"], 4), pct(result["wacc_foreign"], 4),
            ],
        }
    )
    st.dataframe(results_table, hide_index=True, width="stretch")
    st.caption(
        f"Shared by both currencies: unlevered beta {beta_u:.4f}, tax rate {tax_pct:.2f}% "
        f"({tax_choice}), debt share {debt_pct}%, equity risk premium {erp_pct:.2f}%. "
        f"Conversion factor (1 + {rf_f_pct:.2f}%) / (1 + {rf_us_pct:.2f}%) = {result['factor']:.6f}."
    )

    # ---- sources and dates -------------------------------------------------
    def status(live: bool, manual: bool, fetched: str) -> str:
        if manual:
            return "Manual override"
        return f"Live, fetched {fetched}" if live else "Snapshot (live download failed)"

    te_us = f"[Trading Economics]({TE_COUNTRY_URL.format(slug=COUNTRIES[US][1])})"
    te_f = f"[Trading Economics]({TE_COUNTRY_URL.format(slug=COUNTRIES[code][1])})"
    oecd = f"[OECD Corporate Tax Statistics]({OECD_PAGE_URL}), combined statutory rate"
    source_rows = [
        (f"Unlevered beta: {industry}", f"{beta_u:.4f}",
         f"[Damodaran, U.S. industry betas]({BETAS_URL})", beta_data["date"],
         status(beta_data["ok"], beta_manual is not None, beta_data["fetched"])),
        ("U.S. corporate tax rate", f"{tax_us_pct:.2f}%", oecd, tax_us_year,
         status(tax_us_live, tax_us_manual is not None, tax_data["fetched"])),
    ]
    if not is_us:
        source_rows.append(
            (f"{country} corporate tax rate", f"{tax_f_pct:.2f}%", oecd, tax_f_year,
             status(tax_f_live, tax_f_manual is not None, tax_data["fetched"])))
    source_rows += [
        ("USD BBB spread (Baa2/BBB, large non-financial firms)", f"{spread_pct:.2f}%",
         f"[Damodaran, ratings and spreads]({RATINGS_URL})", spread_data["date"],
         status(spread_data["ok"], spread_manual is not None, spread_data["fetched"])),
        ("U.S. 10-year government yield", f"{rf_us_pct:.2f}%", te_us, yield_us["date"],
         status(yield_us["ok"], rf_us_manual is not None, yield_us["fetched"])),
    ]
    if not is_us:
        source_rows.append(
            (f"{country} 10-year government yield", f"{rf_f_pct:.2f}%", te_f,
             yield_foreign["date"],
             status(yield_foreign["ok"], rf_f_manual is not None, yield_foreign["fetched"])))
    source_rows += [
        (f"Foreign-currency BBB spread ({currency})", f"{result['spread_foreign'] * 100:.4f}%",
         "USD spread × (1 + Rf*) / (1 + Rf)" if spread_f_manual is None else "User input", "",
         "Automatic" if spread_f_manual is None else "Manual override"),
        ("Equity risk premium", f"{erp_pct:.2f}%", "User assumption", "", "User input"),
        ("Target debt share", f"{debt_pct}%", "User assumption", "", "User input"),
    ]
    st.subheader("Inputs, sources and dates")
    lines = ["| Input | Value used | Source | Data date | Status |", "|---|---|---|---|---|"]
    lines += [f"| {a} | {b} | {c} | {d if b and not e.startswith('Manual') else ''} | {e} |"
              for a, b, c, d, e in source_rows]
    st.markdown("\n".join(lines))

    # ---- step by step ------------------------------------------------------
    with st.expander("Step-by-step calculation"):
        d, t = inputs["debt_share"], inputs["tax"]
        steps = [
            ("D/E = d / (1 − d)", f"{d:.2f} / {1 - d:.2f}", f"{result['debt_to_equity']:.6f}"),
            ("βE = βU × [1 + (1 − τ) × D/E]",
             f"{beta_u:.6f} × [1 + {1 - t:.6f} × {result['debt_to_equity']:.6f}]",
             f"{result['beta_e']:.6f}"),
            ("Re = Rf + βE × ERP",
             f"{rf_us_pct:.4f}% + {result['beta_e']:.6f} × {erp_pct:.2f}%", pct(result["re_usd"], 4)),
            ("Rd = Rf + s", f"{rf_us_pct:.4f}% + {spread_pct:.4f}%", pct(result["rd_usd"], 4)),
            ("Rd × (1 − τ)", f"{pct(result['rd_usd'], 4)} × {1 - t:.6f}",
             pct(result["rd_usd_after_tax"], 4)),
            ("WACC = (1 − d) × Re + d × Rd × (1 − τ)",
             f"{1 - d:.2f} × {pct(result['re_usd'], 4)} + {d:.2f} × {pct(result['rd_usd_after_tax'], 4)}",
             pct(result["wacc_usd"], 4)),
            ("F = (1 + Rf*) / (1 + Rf)", f"(1 + {rf_f_pct:.4f}%) / (1 + {rf_us_pct:.4f}%)",
             f"{result['factor']:.6f}"),
            ("Re* = (1 + Re) × F − 1", f"(1 + {pct(result['re_usd'], 4)}) × {result['factor']:.6f} − 1",
             pct(result["re_foreign"], 4)),
            ("s* = s × F" if spread_f_manual is None else "s* (manual override)",
             f"{spread_pct:.4f}% × {result['factor']:.6f}" if spread_f_manual is None else "user input",
             pct(result["spread_foreign"], 4)),
            ("Rd* = Rf* + s*", f"{rf_f_pct:.4f}% + {pct(result['spread_foreign'], 4)}",
             pct(result["rd_foreign"], 4)),
            ("Rd* × (1 − τ)", f"{pct(result['rd_foreign'], 4)} × {1 - t:.6f}",
             pct(result["rd_foreign_after_tax"], 4)),
            ("WACC* = (1 − d) × Re* + d × Rd* × (1 − τ)",
             f"{1 - d:.2f} × {pct(result['re_foreign'], 4)} + {d:.2f} × {pct(result['rd_foreign_after_tax'], 4)}",
             pct(result["wacc_foreign"], 4)),
        ]
        st.markdown(
            "| Formula | With the numbers | Result |\n|---|---|---|\n"
            + "\n".join(f"| {a} | {b} | {c} |" for a, b, c in steps)
        )
        st.caption("An asterisk marks the foreign currency. Calculations use full precision; "
                   "displayed values are rounded.")

    # ---- all-industry comparison -------------------------------------------
    st.subheader("All-industry comparison")
    all_betas = pd.Series(beta_data["betas"], name="Unlevered beta")
    everything = cost_of_capital(all_betas, **inputs)
    comparison = pd.DataFrame(
        {
            "Industry": all_betas.index,
            "Unlevered beta": all_betas.to_numpy(),
            "Relevered equity beta": everything["beta_e"].to_numpy(),
            "Cost of equity, USD": everything["re_usd"].to_numpy(),
            "WACC, USD": everything["wacc_usd"].to_numpy(),
            f"Cost of equity, {foreign_label}": everything["re_foreign"].to_numpy(),
            f"WACC, {foreign_label}": everything["wacc_foreign"].to_numpy(),
        }
    )
    st.caption(
        f"All {len(comparison)} Damodaran industries with the same yields, spread, tax rate, "
        "debt share and equity risk premium as above. Each industry uses its downloaded "
        "unlevered beta; a manual beta override applies only to the selected project."
    )
    if is_us:
        chart_column, chart_label = "WACC, USD", "WACC in USD"
    else:
        pick = st.radio("Chart currency", ["USD", foreign_label], horizontal=True)
        chart_column = "WACC, USD" if pick == "USD" else f"WACC, {foreign_label}"
        chart_label = f"WACC in {pick}"
    tab_chart, tab_table = st.tabs(["Chart", "Table"])
    with tab_chart:
        with st.container(height=560):
            st.plotly_chart(industry_chart(comparison, chart_column, chart_label, industry),
                            width="stretch", config={"displayModeBar": False})
    with tab_table:
        shown = comparison.copy()
        for name in shown.columns[3:]:
            shown[name] = shown[name] * 100
        st.dataframe(
            shown, hide_index=True, width="stretch", height=520,
            column_config={
                "Unlevered beta": st.column_config.NumberColumn(format="%.3f"),
                "Relevered equity beta": st.column_config.NumberColumn(format="%.3f"),
                **{name: st.column_config.NumberColumn(format="%.2f%%") for name in shown.columns[3:]},
            },
        )

    # ---- downloads ---------------------------------------------------------
    project = pd.DataFrame(
        [
            ("Industry", industry, ""), ("Target country", country, ""), ("Currency", "USD", currency),
            ("Unlevered beta", beta_u, beta_u),
            ("Project tax rate basis", tax_choice, tax_choice),
            ("Project tax rate", inputs["tax"], inputs["tax"]),
            ("Equity risk premium", inputs["erp"], inputs["erp"]),
            ("Target debt share", d, d),
            ("Debt-to-equity ratio", result["debt_to_equity"], result["debt_to_equity"]),
            ("Relevered equity beta", result["beta_e"], result["beta_e"]),
            ("10-year government yield", inputs["rf_us"], inputs["rf_foreign"]),
            ("BBB borrowing spread", inputs["spread_us"], result["spread_foreign"]),
            ("Foreign spread basis", "", spread_f_note),
            ("Conversion factor (1+Rf*)/(1+Rf)", 1.0, result["factor"]),
            ("Cost of equity", result["re_usd"], result["re_foreign"]),
            ("Pretax cost of debt", result["rd_usd"], result["rd_foreign"]),
            ("After-tax cost of debt", result["rd_usd_after_tax"], result["rd_foreign_after_tax"]),
            ("WACC", result["wacc_usd"], result["wacc_foreign"]),
        ],
        columns=["Item", "USD", foreign_label if not is_us else "USD (selected country)"],
    )
    project = project.map(lambda v: round(v, 10) if isinstance(v, float) else v)
    one, two = st.columns(2)
    one.download_button(
        "Download project results (CSV)", project.to_csv(index=False).encode("utf-8"),
        file_name=f"cost_of_capital_{code}.csv", mime="text/csv",
    )
    two.download_button(
        "Download all-industry comparison (CSV)", comparison.to_csv(index=False, float_format="%.10g").encode("utf-8"),
        file_name=f"cost_of_capital_all_industries_{code}.csv", mime="text/csv",
    )
    st.caption("Rates in the CSV files are decimals (0.085 means 8.5%).")

    # ---- assumptions -------------------------------------------------------
    st.subheader("Main assumptions")
    st.markdown(
        f"""
- **Risk benchmark.** Risk is measured against the U.S. equity market with one equity risk premium ({erp_pct:.2f}%). There is no separate country risk premium.
- **Project risk.** The unlevered beta of the U.S. industry stands in for the project's business risk, wherever the project is located.
- **Relevering.** The beta is relevered at the project's own debt share and tax rate, assuming debt and the interest tax shield have zero beta.
- **Risk-free rates.** The 10-year government yields are used as risk-free rates. They also contain term and liquidity premia and, for some countries, sovereign default risk, so two countries that share a currency can give different results.
- **Currency conversion.** The USD cost of equity is converted with (1 + Rf*) / (1 + Rf). This assumes equal real interest rates and relative purchasing power parity, so the yield gap equals expected inflation and expected depreciation. It adds no premium for currency risk.
- **Borrowing cost.** The USD cost of debt is the U.S. yield plus the BBB spread. Unless overridden, the foreign spread is the USD spread multiplied by the same conversion factor.
- **Taxes and financing.** One tax rate ({tax_choice}, {tax_pct:.2f}%), one debt share and one equity risk premium apply in both currencies. The debt share is a constant target, and interest is fully deductible when paid.
- **Equal yields.** When both yields are equal and the foreign spread override is off, the results in the two currencies are identical.
"""
    )


if __name__ == "__main__":
    main()
