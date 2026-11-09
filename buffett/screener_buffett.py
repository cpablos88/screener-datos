# %% [markdown]
# # Screener Buffett — S&P 500 (v2)
#
# Calcula cada día la puntuación Buffett (0–100) de todas las empresas del S&P 500
# siguiendo los 5 pasos y los 4 factores de Warren Buffett.
#
# Funciona igual en **GitHub Actions** (recomendado: corre solo cada día laborable) y en **Google Colab**.
#
# Salidas (en `OUTPUT_DIR`):
# - `screener_buffett.csv`         → screener completo del último cierre (todas las columnas)
# - `screener_buffett_agente.csv`  → versión compacta para el agente (sin las DESCARTAR)
# - `mercado.json`                 → resumen del día (bono, tasa usada, señales, mediana Precio/VI)
# - `cambios.csv`                  → cambios de señal respecto al día anterior
# - `historico/AAAA-MM-DD.csv`     → copia compacta de cada día
# - `screener_buffett_AAAA-MM-DD.xlsx` → solo en Colab (WRITE_XLSX)
#
# **Fuentes:** SEC EDGAR (10–12 años de estados financieros, API pública "frames")
# y Yahoo Finance vía `yfinance` (precio, capitalización, analistas, fecha de resultados).
#
# **Identificación SEC:** la SEC exige nombre y email. En GitHub va como secreto `SEC_USER_AGENT`
# (nunca en el código, porque el repositorio es público). En Colab, escríbelo en CONFIG.

# %%
# !pip -q install yfinance openpyxl lxml   # ← descomenta en Colab

# %%
import os as _os

CONFIG = {
    # --- Identificación para la SEC (obligatorio). En GitHub: secreto SEC_USER_AGENT ---
    "SEC_USER_AGENT": (_os.environ.get("SEC_USER_AGENT") or "").strip() or "Nombre Apellido tu_email@ejemplo.com",

    # --- Carpetas ---
    "OUTPUT_DIR": _os.environ.get("BUFFETT_OUTPUT_DIR", "/content/drive/MyDrive/ScreenerBuffett"),
    "CACHE_DIR": _os.environ.get("BUFFETT_CACHE_DIR", ""),        # vacío → OUTPUT_DIR/cache
    "REVIEWS_FILE": _os.environ.get("BUFFETT_REVIEWS", ""),       # vacío → OUTPUT_DIR/revisiones.csv
    "WRITE_XLSX": _os.environ.get("BUFFETT_XLSX", "1") == "1",    # en GitHub se pone a 0
    "FORCE": _os.environ.get("FORZAR", "0") == "1",               # recalcular aunque la sesión ya esté hecha
    "PREWARM": _os.environ.get("BUFFETT_PREWARM", "0") == "1",    # solo precargar SEC, histórico y analistas (tarde)

    # --- Caché ---
    "CACHE_DAYS_FUNDAMENTALS": 3,     # refresco de los datos SEC de los 2 últimos años
    "CACHE_DAYS_HISTORY": 7,          # refresco del histórico mensual de precios

    # --- Valoración (Paso 4) ---
    "DISCOUNT_RATE": 0.10,            # suelo de la tasa de descuento
    "DISCOUNT_SPREAD": 0.05,          # tasa = máx(DISCOUNT_RATE, bono 10a + DISCOUNT_SPREAD)
    "TERMINAL_GROWTH": 0.03,          # crecimiento perpetuo a partir del año 10
    "MAX_GROWTH": 0.12,               # tope al crecimiento usado en los años 1–5 (prudencia Buffett)
    "PE_TERMINAL_CAP": 20.0,          # PER máximo de salida en el método de TAE a 10 años
    "PE_TERMINAL_FLOOR": 8.0,
    "TARGET_RETURN": 0.15,            # rentabilidad para "Precio para 15%"
    "USE_ANALYST_GROWTH": True,       # usa el crecimiento a 5 años de analistas solo si es MENOR que el histórico
    "CYCLICAL_DROPS": 2,              # nº de caídas de ventas >10% a partir del cual se normaliza el beneficio

    # --- Margen de seguridad y señales (Paso 5) ---
    "QUALITY_WONDERFUL": 52.0,        # calidad (F1+F2+F3, sobre 70) para "Maravillosa"
    "QUALITY_GOOD": 42.0,             # calidad mínima para "Buena"; por debajo → DESCARTAR
    "MS_WONDERFUL": 0.20,             # margen de seguridad exigido a una empresa maravillosa
    "MS_GOOD": 0.33,                  # margen de seguridad exigido a una empresa buena
    "MIN_TAE_BUY": 0.10,              # TAE a 10 años mínima para COMPRAR
    "WATCH_BAND": 0.15,               # VIGILAR si el precio está hasta un 15% por encima del precio de compra
    "OVERVALUED": 1.50,               # SOBREVALORADA si precio ≥ 1,5 × valor intrínseco
    "ANOMALY_MS": 0.60,               # COMPRAR con MS > 60% → REVISAR (verificar datos)
    "ANOMALY_OE_YIELD": 0.25,         # COMPRAR con OE yield > 25% → REVISAR
    "EROSION_ROIC_PP": 5.0,           # alerta si el ROIC cae ≥ 5 puntos (media 3 últimos vs 3 primeros)
    "EROSION_GM_PP": 3.0,             # alerta si el margen bruto cae ≥ 3 puntos
    "REVIEW_MAX_DAYS": 90,            # una revisión cualitativa caduca a los 90 días

    # --- Datos ---
    "UST10_FALLBACK": 0.0425,         # rentabilidad del bono a 10 años si Yahoo falla
    "MAX_WORKERS": int(_os.environ.get("BUFFETT_WORKERS", "1")),  # 1 = secuencial con pausas (lo más seguro con Yahoo)
    "YAHOO_PAUSE": 0.4,               # pausa entre empresas en modo secuencial (s)
    "LIMIT": int(_os.environ["BUFFETT_LIMIT"]) if _os.environ.get("BUFFETT_LIMIT") else None,
}

# %%
import datetime as dt
import hashlib
import io
import json
import math
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests

TODAY = dt.date.today()
NY = ZoneInfo("America/New_York")


def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def _ok(x):
    return x is not None and isinstance(x, (int, float, np.floating, np.integer)) and math.isfinite(float(x))


def _fresh(path, max_age_days):
    if not os.path.exists(path):
        return False
    age = (time.time() - os.path.getmtime(path)) / 86400
    return age <= max_age_days


def _pct(x, nd=1):
    return round(float(x) * 100, nd) if _ok(x) else None


def _rnd(x, nd=2):
    return round(float(x), nd) if _ok(x) else None


def cache_dir(cfg):
    d = cfg["CACHE_DIR"] or os.path.join(cfg["OUTPUT_DIR"], "cache")
    os.makedirs(d, exist_ok=True)
    return d


# =====================================================================
# 1. UNIVERSO S&P 500
# =====================================================================
WIKI_URL = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
FALLBACK_URL = "https://raw.githubusercontent.com/datasets/s-and-p-500-companies/main/data/constituents.csv"


def get_universe(cfg):
    cache = os.path.join(cache_dir(cfg), "universe.csv")
    headers = {"User-Agent": "Mozilla/5.0 (screener-buffett; educational)"}
    df = None
    for src in ("wiki", "github"):
        try:
            if src == "wiki":
                html = requests.get(WIKI_URL, headers=headers, timeout=30).text
                t = pd.read_html(io.StringIO(html), attrs={"id": "constituents"})[0]
            else:
                t = pd.read_csv(FALLBACK_URL)
            df = pd.DataFrame({
                "ticker": t["Symbol"].astype(str).str.strip(),
                "empresa": t["Security"].astype(str),
                "sector": t["GICS Sector"].astype(str),
                "subsector": t["GICS Sub-Industry"].astype(str),
                "cik": pd.to_numeric(t["CIK"], errors="coerce"),
            })
            break
        except Exception as e:  # noqa: BLE001
            log(f"Lista S&P 500 ({src}) no disponible: {e}")
    if df is None or df.empty:
        if os.path.exists(cache):
            log("Uso la última lista guardada en caché.")
            return pd.read_csv(cache)
        raise RuntimeError("No se pudo obtener la lista del S&P 500.")
    df["yf"] = df["ticker"].str.replace(".", "-", regex=False)
    df = df.dropna(subset=["cik"]).copy()
    df["cik"] = df["cik"].astype(int)
    df.to_csv(cache, index=False)
    return df


# =====================================================================
# 2. FUNDAMENTALES 10–12 AÑOS — SEC EDGAR (API "frames")
# =====================================================================
DURATION_TAGS = {
    "revenue": ("USD", ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
                        "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet",
                        "RevenuesNetOfInterestExpense", "SalesRevenueGoodsNet"]),
    "net_income": ("USD", ["NetIncomeLoss", "NetIncomeLossAvailableToCommonStockholdersBasic", "ProfitLoss"]),
    "gross_profit": ("USD", ["GrossProfit"]),
    "cogs": ("USD", ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold"]),
    "op_income": ("USD", ["OperatingIncomeLoss"]),
    "pretax": ("USD", ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                       "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"]),
    "tax": ("USD", ["IncomeTaxExpenseBenefit"]),
    "cfo": ("USD", ["NetCashProvidedByUsedInOperatingActivities",
                    "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"]),
    "capex": ("USD", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
                      "PaymentsForCapitalImprovements"]),
    "sbc": ("USD", ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"]),
    "dividends": ("USD", ["PaymentsOfDividendsCommonStock", "PaymentsOfDividends"]),
    "shares": ("shares", ["WeightedAverageNumberOfDilutedSharesOutstanding",
                          "WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
                          "WeightedAverageNumberOfSharesOutstandingBasic"]),
}
INSTANT_TAGS = {
    "equity": ("USD", ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    "assets": ("USD", ["Assets"]),
    "lt_debt": ("USD", ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations", "LongTermDebt"]),
    "cash": ("USD", ["CashAndCashEquivalentsAtCarryingValue",
                     "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"]),
    "st_inv": ("USD", ["ShortTermInvestments", "MarketableSecuritiesCurrent",
                       "AvailableForSaleSecuritiesDebtSecuritiesCurrent"]),
}
SEC_FRAMES = "https://data.sec.gov/api/xbrl/frames"


class SecFrames:
    def __init__(self, cfg):
        self.cfg = cfg
        self.dir = os.path.join(cache_dir(cfg), "sec_frames")
        os.makedirs(self.dir, exist_ok=True)
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": cfg["SEC_USER_AGENT"], "Accept-Encoding": "gzip, deflate"})
        self.calls = 0
        self.ok = 0
        self.forbidden = 0

    def get(self, tag, unit, period, year):
        fn = os.path.join(self.dir, f"{tag}_{unit}_{period}.json")
        old_year = year <= TODAY.year - 3
        max_age = 3650 if old_year else self.cfg["CACHE_DAYS_FUNDAMENTALS"]
        if _fresh(fn, max_age):
            with open(fn) as f:
                return json.load(f)
        url = f"{SEC_FRAMES}/us-gaap/{tag}/{unit}/{period}.json"
        for attempt in range(4):
            try:
                r = self.s.get(url, timeout=60)
                self.calls += 1
                time.sleep(0.12)  # límite SEC: 10 peticiones/s
                if r.status_code == 403:
                    self.forbidden += 1
                    if self.ok == 0 and self.forbidden >= 3:
                        raise RuntimeError("La SEC rechaza las descargas (403). Revisa el secreto SEC_USER_AGENT: "
                                           "debe ser «Nombre Apellido email» con un email real.")
                    time.sleep(2)
                    continue
                if r.status_code in (200, 404):
                    self.ok += 1
                    data = {"data": [{"cik": d["cik"], "val": d["val"]} for d in r.json().get("data", [])]} \
                        if r.status_code == 200 else {"data": []}
                    with open(fn, "w") as f:
                        json.dump(data, f)
                    return data
            except (requests.RequestException, ValueError):
                pass
            time.sleep(2 * (attempt + 1))
        log(f"  SEC sin respuesta: {tag} {period}")
        if os.path.exists(fn):  # mejor un dato de hace días que ninguno
            with open(fn) as f:
                return json.load(f)
        return {"data": []}


def build_fundamentals(ciks, years, fetch):
    """fetch(tag, unit, period, year) -> {"data":[{"cik":..,"val":..}]}.
    Devuelve {cik: DataFrame(index=año, columnas=métricas)}."""
    ciks = set(int(c) for c in ciks)
    store = {}

    def collect(metric, unit, tags, period_fmt):
        for y in years:
            got = {}
            for tag in tags:
                if len(got) >= len(ciks):
                    break
                data = fetch(tag, unit, period_fmt.format(y=y), y)
                for d in data.get("data", []):
                    c = int(d["cik"])
                    if c in ciks and c not in got and d.get("val") is not None:
                        got[c] = float(d["val"])
            for c, v in got.items():
                store.setdefault(c, {}).setdefault(y, {})[metric] = v

    for metric, (unit, tags) in DURATION_TAGS.items():
        log(f"  SEC: {metric}")
        collect(metric, unit, tags, "CY{y}")
    for metric, (unit, tags) in INSTANT_TAGS.items():
        log(f"  SEC: {metric}")
        collect(metric, unit, tags, "CY{y}Q4I")

    panels = {}
    for c, by_year in store.items():
        df = pd.DataFrame.from_dict(by_year, orient="index").sort_index()
        df.index = df.index.astype(int)
        panels[c] = df
    return panels


def yf_statements(yf_ticker):
    """Plan B si la SEC da menos de 5 años: estados de Yahoo (4–5 años)."""
    import yfinance as yf
    t = yf.Ticker(yf_ticker)
    inc, bal, cf = t.income_stmt, t.balance_sheet, t.cashflow

    def row(df, names, sign=1):
        for n in names:
            if df is not None and n in df.index:
                return df.loc[n] * sign
        return None

    cols = {
        "revenue": row(inc, ["Total Revenue", "Operating Revenue"]),
        "net_income": row(inc, ["Net Income Common Stockholders", "Net Income"]),
        "gross_profit": row(inc, ["Gross Profit"]),
        "cogs": row(inc, ["Cost Of Revenue"]),
        "op_income": row(inc, ["Operating Income"]),
        "pretax": row(inc, ["Pretax Income"]),
        "tax": row(inc, ["Tax Provision"]),
        "shares": row(inc, ["Diluted Average Shares", "Basic Average Shares"]),
        "equity": row(bal, ["Stockholders Equity", "Common Stock Equity"]),
        "assets": row(bal, ["Total Assets"]),
        "lt_debt": row(bal, ["Long Term Debt", "Long Term Debt And Capital Lease Obligation"]),
        "cash": row(bal, ["Cash And Cash Equivalents"]),
        "st_inv": row(bal, ["Other Short Term Investments"]),
        "cfo": row(cf, ["Operating Cash Flow"]),
        "capex": row(cf, ["Capital Expenditure"], -1),
        "sbc": row(cf, ["Stock Based Compensation"]),
        "dividends": row(cf, ["Cash Dividends Paid", "Common Stock Dividend Paid"], -1),
    }
    cols = {k: v for k, v in cols.items() if v is not None}
    if not cols:
        return None
    df = pd.DataFrame(cols)
    df.index = [d.year if d.month >= 7 else d.year - 1 for d in pd.to_datetime(df.index)]
    df = df[~df.index.duplicated(keep="first")].sort_index()
    return df.astype(float)


# =====================================================================
# 3. MERCADO — Yahoo Finance
# =====================================================================
def _drop_open_session(df):
    """Quita la vela de hoy si Wall Street aún no ha cerrado (16:20 Nueva York)."""
    now = dt.datetime.now(NY)
    if len(df) and df.index[-1].date() == now.date() and now.time() < dt.time(16, 20):
        return df.iloc[:-1]
    return df


def get_session():
    import yfinance as yf
    spy = _drop_open_session(yf.download("SPY", period="10d", interval="1d", auto_adjust=False, progress=False))
    return spy.index[-1].date().isoformat()


def get_prices(yf_tickers, cfg):
    import yfinance as yf
    daily = yf.download(yf_tickers, period="10d", interval="1d", auto_adjust=False,
                        progress=False, threads=True)["Close"]
    daily = _drop_open_session(daily).ffill()
    last_date = daily.index[-1].date()
    last = daily.iloc[-1]

    hist_file = os.path.join(cache_dir(cfg), "monthly_history.pkl")
    if _fresh(hist_file, cfg["CACHE_DAYS_HISTORY"]):
        monthly, splits = pd.read_pickle(hist_file)
    else:
        log("  Yahoo: histórico mensual 12 años")
        h = yf.download(yf_tickers, period="12y", interval="1mo", auto_adjust=False, actions=True,
                        progress=False, threads=True)
        monthly = h["Close"]
        splits = h["Stock Splits"] if "Stock Splits" in h.columns.get_level_values(0) else pd.DataFrame()
        pd.to_pickle((monthly, splits), hist_file)
    return last, last_date, monthly, splits


def get_ust10(cfg):
    try:
        import yfinance as yf
        v = float(yf.download("^TNX", period="10d", progress=False, auto_adjust=False)["Close"].dropna().iloc[-1].squeeze())
        return v / 100 if v < 25 else v / 1000
    except Exception:  # noqa: BLE001
        return cfg["UST10_FALLBACK"]


def yf_analyst(yf_ticker, cfg, want_growth):
    """Analistas, capitalización y fecha de resultados (con reintentos)."""
    import yfinance as yf
    out = {}
    for attempt in range(3):
        try:
            t = yf.Ticker(yf_ticker)
            info = t.info or {}
            out.update({
                "market_cap": info.get("marketCap"),
                "mc_price": info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"),
                "shares_out": info.get("impliedSharesOutstanding") or info.get("sharesOutstanding"),
                "trailing_pe": info.get("trailingPE"),
                "target_mean": info.get("targetMeanPrice"),
                "target_low": info.get("targetLowPrice"),
                "target_high": info.get("targetHighPrice"),
                "rec_key": info.get("recommendationKey"),
                "rec_mean": info.get("recommendationMean"),
                "n_analysts": info.get("numberOfAnalystOpinions"),
                "earnings_ts": info.get("earningsTimestampStart") or info.get("earningsTimestamp"),
            })
            try:
                rs = t.recommendations_summary
                if rs is not None and not rs.empty:
                    r0 = rs[rs["period"] == "0m"].iloc[0] if "period" in rs.columns else rs.iloc[0]
                    out["votes"] = [int(r0.get("strongBuy", 0)) + int(r0.get("buy", 0)), int(r0.get("hold", 0)),
                                    int(r0.get("sell", 0)) + int(r0.get("strongSell", 0))]
            except Exception:  # noqa: BLE001
                pass
            if want_growth:
                try:
                    ge = t.growth_estimates
                    for key in ["+5y", "LTG"]:
                        if ge is not None and key in ge.index:
                            v = ge.loc[key].get("stockTrend")
                            if _ok(v):
                                out["growth_5y"] = float(v)
                                break
                except Exception:  # noqa: BLE001
                    pass
            return out
        except Exception:  # noqa: BLE001
            time.sleep(8 * (attempt + 1))
    return out


def get_analysts(yf_tickers, cfg):
    """Analistas del día (UTC). La precarga de la tarde los deja en caché para el cálculo de la noche."""
    cd = cache_dir(cfg)
    cache = os.path.join(cd, f"analysts_{TODAY.isoformat()}.json")
    res = {}
    if os.path.exists(cache):
        with open(cache) as f:
            res = json.load(f)
    week = TODAY.isocalendar()
    gfile = os.path.join(cd, f"growth_{week[0]}W{week[1]:02d}.json")
    growth = {}
    if os.path.exists(gfile):
        with open(gfile) as f:
            growth = json.load(f)
    pending = [t for t in yf_tickers if not res.get(t, {}).get("market_cap")]
    log(f"  Yahoo analistas: {len(pending)} pendientes de {len(yf_tickers)}")

    def one(t):
        r = yf_analyst(t, cfg, want_growth=cfg["USE_ANALYST_GROWTH"] and t not in growth)
        if "growth_5y" in r:
            growth[t] = r["growth_5y"]
        return t, r

    if cfg["MAX_WORKERS"] <= 1:
        for i, t in enumerate(pending, 1):
            k, r = one(t)
            res[k] = r
            time.sleep(cfg["YAHOO_PAUSE"])
            if i % 50 == 0:
                log(f"  Yahoo analistas: {i}/{len(pending)}")
                with open(cache, "w") as f:
                    json.dump(res, f)
    else:
        with ThreadPoolExecutor(max_workers=cfg["MAX_WORKERS"]) as ex:
            for fut in as_completed([ex.submit(one, t) for t in pending]):
                k, r = fut.result()
                res[k] = r
    for t, v in growth.items():
        res.setdefault(t, {})["growth_5y"] = v
    with open(cache, "w") as f:
        json.dump(res, f)
    with open(gfile, "w") as f:
        json.dump(growth, f)
    return res


# =====================================================================
# 4. MÉTRICAS, PUNTUACIÓN Y VALORACIÓN (lógica pura, sin red)
# =====================================================================
FIN_KEYS = ["Bank", "Insurance", "Reinsurance", "Consumer Finance", "Brokerage",
            "Multi-Sector Holdings", "Mortgage"]

CIRCLE_SECTOR = {
    "Consumer Staples": 4, "Financials": 3, "Industrials": 3, "Health Care": 3,
    "Consumer Discretionary": 3, "Communication Services": 3, "Utilities": 3,
    "Real Estate": 2, "Information Technology": 2, "Materials": 2, "Energy": 1,
}
CIRCLE_SUB = [  # (texto contenido en el subsector, puntos) — el primero que coincide manda
    ("Biotechnology", 1), ("Oil & Gas Exploration", 1), ("Oil & Gas Drilling", 1),
    ("Oil & Gas Equipment", 1), ("Oil & Gas Refining", 1), ("Gold", 1), ("Steel", 1), ("Copper", 1),
    ("Aluminum", 1), ("Fertilizers", 1), ("Airlines", 1), ("Commodity Chemicals", 1),
    ("Semiconductor", 2), ("Application Software", 2), ("Technology Hardware", 2),
    ("Electronic Components", 2), ("Communications Equipment", 2), ("Casinos", 2),
    ("Automobile Manufacturers", 2), ("Hotels, Resorts & Cruise Lines", 2), ("Pharmaceuticals", 2),
    ("Homebuilding", 2), ("Integrated Oil & Gas", 2),
    ("Soft Drinks", 4), ("Packaged Foods", 4), ("Household Products", 4), ("Personal Care", 4),
    ("Distillers", 4), ("Brewers", 4), ("Tobacco", 4), ("Consumer Staples Merchandise Retail", 4),
    ("Rail", 4), ("Transaction & Payment Processing", 4), ("Financial Exchanges & Data", 4),
    ("Insurance Brokers", 4), ("Environmental & Facilities Services", 4), ("Home Improvement Retail", 4),
]


def circle_prior(sector, sub):
    for key, pts in CIRCLE_SUB:
        if key in str(sub):
            return pts
    return CIRCLE_SECTOR.get(str(sector), 2)


def is_fin_mode(sub):
    s = str(sub)
    return any(k in s for k in FIN_KEYS) and "Insurance Brokers" not in s


def bracket(x, rules, higher=True):
    """rules: [(umbral, puntos), ...] de mejor a peor. Devuelve None si x no existe."""
    if not _ok(x):
        return None
    for thr, pts in rules:
        if (x >= thr) if higher else (x <= thr):
            return pts
    return 0


def pillar(items, pmax, min_cov=0.6):
    """items: [(puntos|None, máximo)]. Reescala sobre las métricas disponibles si cubren ≥60%."""
    total_max = sum(m for _, m in items)
    avail = [(p, m) for p, m in items if p is not None]
    got = sum(p for p, _ in avail)
    mx = sum(m for _, m in avail)
    if mx == 0:
        return 0.0, 0.0
    cov = mx / total_max
    return (got / mx * pmax if cov >= min_cov else got / total_max * pmax), cov


COMMON_SPLITS = [2, 3, 4, 5, 10, 20, 1.5, 0.5, 1 / 3, 0.25, 0.2, 0.1]


def _split_like(r):
    for c in COMMON_SPLITS:
        if abs(r / c - 1) < 0.12:
            return c
    return None


def adjust_shares(shares, splits_by_year=None):
    """Lleva la serie histórica de acciones a una base homogénea (splits no reexpresados)."""
    s = shares.dropna().astype(float)
    s = s[s > 0].copy()
    yrs = list(s.index)
    adjusted = False
    for i in range(1, len(yrs)):
        y0, y1 = yrs[i - 1], yrs[i]
        r = s[y1] / s[y0]
        if r >= 1.8 or r <= 0.55:
            f = None
            if splits_by_year:
                cand = 1.0
                for yy, ratio in splits_by_year.items():
                    if y0 <= yy <= y1 + 1:
                        cand *= ratio
                if cand != 1.0 and abs(r / cand - 1) < 0.2:
                    f = cand
            f = f or _split_like(r)
            if f:
                s.loc[s.index <= y0] = s.loc[s.index <= y0] * f
                adjusted = True
    return s, adjusted


def cagr(a, b, n):
    if not (_ok(a) and _ok(b)) or a <= 0 or b <= 0 or n <= 0:
        return None
    return (b / a) ** (1 / n) - 1


def trend_pp(series):
    """Media de los 3 últimos años menos media de los 3 primeros, en puntos porcentuales."""
    s = series.dropna()
    if len(s) < 6:
        return None
    return float((s.iloc[-3:].mean() - s.iloc[:3].mean()) * 100)


def growth_path(g, gt):
    return [g if t <= 5 else g + (gt - g) * (t - 5) / 5 for t in range(1, 11)]


def discount_rate(cfg, ust10):
    return max(cfg["DISCOUNT_RATE"], (ust10 or 0) + cfg["DISCOUNT_SPREAD"])


def dcf_value(oe0, g, cfg, r=None):
    r = cfg["DISCOUNT_RATE"] if r is None else r
    gt = cfg["TERMINAL_GROWTH"]
    pv, cf = 0.0, oe0
    for t, gr in enumerate(growth_path(g, gt), 1):
        cf *= 1 + gr
        pv += cf / (1 + r) ** t
    tv = cf * (1 + gt) / (r - gt)
    return pv + tv / (1 + r) ** 10


def compute_company(meta, panel, mkt, ust10, monthly_px, splits_by_year, cfg):
    """meta: dict(ticker, empresa, sector, subsector, fuente). panel: DataFrame anual.
    mkt: dict(price, market_cap, shares_out, trailing_pe, analistas...). Devuelve una fila (dict)."""
    alerts = []
    fin = is_fin_mode(meta["subsector"])
    sub = str(meta["subsector"])
    if "Multi-Sector Holdings" in sub:
        alerts.append("Holding/aseguradora: valoración contable poco fiable, revisar a mano")
    if str(meta["sector"]) == "Real Estate":
        alerts.append("REIT: revisar con FFO/AFFO")

    p = panel.copy() if panel is not None else pd.DataFrame()
    for col in set(list(DURATION_TAGS) + list(INSTANT_TAGS)):
        if col not in p.columns:
            p[col] = np.nan
    p = p[p["net_income"].notna()].sort_index().tail(10)
    # Ejercicios que no cierran en diciembre: el balance del último año aún no existe en la SEC
    # (se lee al 31-dic). Se usa el del año anterior, como máximo un año.
    inst = list(INSTANT_TAGS)
    p[inst] = p[inst].ffill(limit=1)
    price = mkt.get("price")
    row = {"Ticker": meta["ticker"], "Empresa": meta["empresa"], "Sector": meta["sector"],
           "Subsector": meta["subsector"], "Precio": _rnd(price), "Modo": "Financiero" if fin else "General",
           "Fuente": meta.get("fuente", "SEC")}
    n = len(p)
    row["Años historia"] = n
    r_disc = discount_rate(cfg, ust10)

    # ----- Acciones homogéneas y por acción -----
    shares_adj, did_adj = adjust_shares(p["shares"], splits_by_year) if p["shares"].notna().any() else (pd.Series(dtype=float), False)
    if did_adj:
        alerts.append("Splits ajustados en histórico")
    shares_basis = None
    mc_px = mkt.get("mc_price") if _ok(mkt.get("mc_price")) and mkt.get("mc_price") > 0 else price
    if _ok(mkt.get("market_cap")) and _ok(mc_px) and mc_px > 0:
        shares_basis = mkt["market_cap"] / mc_px  # capitalización y precio del mismo momento
    elif _ok(mkt.get("shares_out")):
        shares_basis = float(mkt["shares_out"])
    elif len(shares_adj):
        shares_basis = float(shares_adj.iloc[-1])
    if len(shares_adj) and shares_basis:
        f = shares_basis / shares_adj.iloc[-1]
        if f > 1.4 or f < 0.7:
            shares_adj = shares_adj * f
            alerts.append(f"Base de acciones ajustada ×{f:.3g} (split reciente o varias clases)")
    sh = shares_adj.reindex(p.index)

    ni, rev, eq = p["net_income"], p["revenue"], p["equity"]
    eps = ni / sh
    dps = (p["dividends"].fillna(0)) / sh

    # ----- PASO 1: negocio comprensible y predecible -----
    yrs_profit = int((ni > 0).sum())
    frac_profit = yrs_profit / n if n else 0
    rev_g = rev.pct_change().dropna()
    drops = int((rev_g < -0.10).sum()) if len(rev_g) >= 3 else None
    roe_series = []
    for i, y in enumerate(p.index):
        e_now = eq.loc[y]
        e_prev = eq.iloc[i - 1] if i > 0 else np.nan
        base = np.nanmean([e_now, e_prev]) if _ok(e_prev) and e_prev > 0 else e_now
        if _ok(base) and base > 0 and _ok(ni.loc[y]):
            roe_series.append(ni.loc[y] / base)
    roe = pd.Series(roe_series, dtype=float)
    neg_equity = bool((eq.tail(3) <= 0).any()) if eq.notna().any() else False
    if neg_equity:
        alerts.append("Patrimonio negativo (recompras): ROE no válido, se usa ROIC")

    tax_rate = (p["tax"] / p["pretax"]).clip(0, 0.35).fillna(0.21)
    nopat = p["op_income"] * (1 - tax_rate)
    ic = eq + p["lt_debt"].fillna(0) - p["cash"].fillna(0) - p["st_inv"].fillna(0)
    roic = (nopat / ic.where(ic > 0)).dropna()
    roa = (ni / p["assets"]).dropna()

    stab_series = roic if (neg_equity or len(roe) < 3) else roe
    stab = float(stab_series.std() * 100) if len(stab_series) >= 3 else None
    circle = circle_prior(meta["sector"], meta["subsector"])
    cyclical = (drops is not None and drops >= cfg["CYCLICAL_DROPS"]) or circle <= 1

    f1_items = [
        (bracket(frac_profit, [(1.0, 5), (0.9, 4), (0.8, 3), (0.7, 2)]) if n >= 3 else None, 5),
        (bracket(drops, [(0, 3), (1, 2), (2, 1)], higher=False), 3),
        (bracket(stab, [(5, 3), (10, 2), (15, 1)], higher=False), 3),
        (circle, 4),
    ]
    F1, cov1 = pillar(f1_items, 15)

    # ----- PASO 2: economía del negocio y foso -----
    roe_avg = float(roe.mean()) if len(roe) >= 3 else None
    roic_avg = float(roic.mean()) if len(roic) >= 3 else None
    roa_avg = float(roa.mean()) if len(roa) >= 3 else None
    lev = (p["assets"] / eq).iloc[-1] if eq.notna().any() and _ok(eq.iloc[-1]) and eq.iloc[-1] > 0 else None
    gp = p["gross_profit"].where(p["gross_profit"].notna(), rev - p["cogs"])
    gm_series = (gp / rev).dropna()
    gm_avg = float(gm_series.mean()) if len(gm_series) >= 3 else None
    nm_avg = float((ni / rev).dropna().mean()) if (ni / rev).notna().sum() >= 3 else None

    # Erosión del foso: tendencia de ROIC (o ROA en financieras) y de margen bruto
    roic_trend = trend_pp(roa if fin else roic)
    gm_trend = None if fin else trend_pp(gm_series)
    erosion_roic = _ok(roic_trend) and roic_trend <= -cfg["EROSION_ROIC_PP"] * (0.2 if fin else 1)
    erosion_gm = _ok(gm_trend) and gm_trend <= -cfg["EROSION_GM_PP"]
    if erosion_roic or erosion_gm:
        alerts.append("Posible erosión del foso (" + ", ".join(
            x for x, ok in (("ROIC/ROA a la baja", erosion_roic), ("margen bruto a la baja", erosion_gm)) if ok) + ")")

    eps_v = eps.dropna()
    eps_cagr = None
    if len(eps_v) >= 5:
        k = 3 if len(eps_v) >= 8 else 2
        a, b = eps_v.iloc[:k].mean(), eps_v.iloc[-k:].mean()
        span = (eps_v.index[-k:].values.mean() - eps_v.index[:k].values.mean())
        eps_cagr = cagr(a, b, span)
    rev_cagr = None
    rv = rev.dropna()
    if len(rv) >= 5:
        rev_cagr = cagr(rv.iloc[:2].mean(), rv.iloc[-2:].mean(), rv.index[-2:].values.mean() - rv.index[:2].values.mean())

    capex_ni = (p["capex"].sum() / ni.sum()) if (p["capex"].notna().sum() >= 3 and ni.sum() > 0) else None
    fcf = p["cfo"] - p["capex"].fillna(0)
    fcf_conv = (fcf.sum() / ni.sum()) if (p["cfo"].notna().sum() >= 3 and ni.sum() > 0) else None

    if fin:
        roe_pts = bracket(roe_avg, [(0.20, 8), (0.15, 6), (0.12, 3)])
        f2_items = [
            (roe_pts, 8),
            (bracket(roa_avg, [(0.012, 7), (0.010, 5), (0.008, 3)]), 7),
            (None, 4),
            (bracket(nm_avg, [(0.20, 4), (0.10, 3), (0.05, 1)]), 4),
            (bracket(eps_cagr, [(0.10, 5), (0.07, 4), (0.04, 2), (0.0001, 1)]), 5),
            (None, 4), (None, 3),
        ]
    else:
        if neg_equity or roe_avg is None:
            roe_pts = bracket(roic_avg, [(0.15, 8), (0.12, 6), (0.10, 3)])
        else:
            roe_pts = bracket(roe_avg, [(0.20, 8), (0.15, 6), (0.12, 3)])
            if roe_pts and _ok(lev) and lev > 5:
                roe_pts = {8: 6, 6: 3, 3: 0}.get(roe_pts, 0)
                alerts.append("ROE apoyado en apalancamiento (activo/patrimonio > 5)")
        f2_items = [
            (roe_pts, 8),
            (bracket(roic_avg, [(0.15, 7), (0.12, 5), (0.10, 3)]), 7),
            (bracket(gm_avg, [(0.40, 4), (0.30, 2)]), 4),
            (bracket(nm_avg, [(0.20, 4), (0.10, 3), (0.05, 1)]), 4),
            (bracket(eps_cagr, [(0.10, 5), (0.07, 4), (0.04, 2), (0.0001, 1)]), 5),
            (bracket(capex_ni, [(0.25, 4), (0.50, 3), (0.75, 1)], higher=False), 4),
            (bracket(fcf_conv, [(0.90, 3), (0.70, 2), (0.50, 1)]), 3),
        ]
    F2, cov2 = pillar(f2_items, 35)

    # ----- PASO 3: dirección honesta y competente -----
    retained_ratio = None
    if len(eps_v) >= 5 and monthly_px is not None and len(monthly_px.dropna()):
        start_y = eps_v.index[0]
        yr_px = monthly_px[monthly_px.index.year == start_y].dropna()
        if len(yr_px) and _ok(price):
            p0 = float(yr_px.mean())  # precio medio del año inicial: menos ruido que un solo cierre
            later = [y for y in eps_v.index if y > start_y]
            retained = float((eps.reindex(later) - dps.reindex(later).fillna(0)).sum())
            if retained > 0:
                retained_ratio = (price - p0) / retained
    shv = shares_adj.dropna()
    sh_cagr = cagr(shv.iloc[0], shv.iloc[-1], shv.index[-1] - shv.index[0]) if len(shv) >= 4 else None
    ni3 = ni.tail(3).mean()
    debt_last = p["lt_debt"].iloc[-1] if len(p) else np.nan
    if not _ok(debt_last) and p["assets"].notna().any() and not fin:
        debt_last = 0.0
        alerts.append("Deuda LP no reportada en XBRL: se asume 0, verificar")
    if not _ok(debt_last):
        debt_years = None
    elif debt_last <= 0:
        debt_years = 0.0
    elif _ok(ni3) and ni3 > 0:
        debt_years = debt_last / ni3
    else:
        debt_years = 99.0
    sbc3 = p["sbc"].tail(3).mean()
    fcf3 = fcf.tail(3).mean()
    sbc_fcf = (sbc3 / fcf3) if (_ok(sbc3) and _ok(fcf3) and fcf3 > 0) else None
    accr = None
    if _ok(ni.iloc[-1] if len(ni) else None) and _ok(p["cfo"].iloc[-1] if len(p) else None) and _ok(p["assets"].iloc[-1] if len(p) else None):
        accr = (ni.iloc[-1] - p["cfo"].iloc[-1]) / p["assets"].iloc[-1]
    eq_assets = (eq / p["assets"]).dropna().iloc[-1] if (eq / p["assets"]).notna().any() else None

    f3_items = [
        (bracket(retained_ratio, [(1.5, 6), (1.0, 4), (0.5, 1)]), 6),
        (bracket(sh_cagr, [(-0.01, 5), (0.005, 4), (0.02, 2)], higher=False), 5),
        (bracket(eq_assets, [(0.10, 5), (0.08, 4), (0.06, 2)]) if fin
         else bracket(debt_years, [(2, 5), (3, 4), (4, 2)], higher=False), 5),
        (bracket((sbc3 / ni3) if (fin and _ok(sbc3) and _ok(ni3) and ni3 > 0) else sbc_fcf,
                 [(0.05, 2), (0.10, 1)], higher=False), 2),
        (None if fin else bracket(accr, [(0.0, 2), (0.05, 1)], higher=False), 2),
    ]
    F3, cov3 = pillar(f3_items, 20)

    # ----- PASO 4: valor intrínseco (owner earnings descontadas) -----
    if fin:
        oe_series = ni
    else:
        oe_series = p["cfo"] - p["capex"].fillna(0) - p["sbc"].fillna(0)
        if p["sbc"].tail(3).isna().all():
            alerts.append("SBC no disponible: owner earnings sin restar SBC")
    oe_v = oe_series.dropna()
    oe3 = oe_v.tail(3).mean() if len(oe_v) else None
    if cyclical and len(oe_v) >= 5 and _ok(oe3):
        oe_cycle = oe_v.mean()  # media del ciclo completo (hasta 10 años)
        if oe_cycle < oe3:
            oe3 = oe_cycle
            alerts.append("Cíclica: beneficios normalizados a la media del ciclo")
    oe_ps = (oe3 / shares_basis) if (_ok(oe3) and shares_basis) else None

    g_cands = [x for x in [eps_cagr, mkt.get("growth_5y")] if _ok(x)]
    if g_cands:
        g = min(g_cands)
    elif _ok(rev_cagr):
        g = rev_cagr
        alerts.append("Crecimiento por ventas (BPA no comparable)")
    else:
        g = 0.0
        alerts.append("Sin crecimiento estimable: g = 0")
    g = float(np.clip(g, 0.0, cfg["MAX_GROWTH"]))

    iv = None
    if _ok(oe_ps) and oe_ps > 0:
        iv = dcf_value(oe_ps, g, cfg, r_disc)
        if not fin and shares_basis:
            net_cash = (p["cash"].fillna(0).iloc[-1] + p["st_inv"].fillna(0).iloc[-1] - (debt_last if _ok(debt_last) else 0))
            if net_cash > 0:
                iv += net_cash / shares_basis

    # TAE esperada a 10 años (método BPA × PER de salida + dividendos)
    tae, price_15 = None, None
    eps_base = eps_v.iloc[-1] if len(eps_v) else None
    if cyclical and len(eps_v) >= 5 and _ok(eps_base):
        eps_base = min(eps_base, float(eps_v.mean()))
    if _ok(eps_base) and eps_base > 0 and _ok(price) and price > 0:
        pes = []
        if monthly_px is not None:
            for y in eps_v.index:
                dec = monthly_px[(monthly_px.index.year == y) & (monthly_px.index.month == 12)].dropna()
                if len(dec) and eps_v.loc[y] > 0:
                    pes.append(float(dec.iloc[-1]) / eps_v.loc[y])
        pe_hist = float(np.median(pes)) if len(pes) >= 3 else (mkt.get("trailing_pe") if _ok(mkt.get("trailing_pe")) else price / eps_base)
        pe_term = float(np.clip(pe_hist, cfg["PE_TERMINAL_FLOOR"], cfg["PE_TERMINAL_CAP"]))
        payout = float(np.clip((dps / eps).tail(3).mean(), 0, 1)) if (dps / eps).tail(3).notna().any() else 0.0
        e, divs = eps_base, 0.0
        for gr in growth_path(g, cfg["TERMINAL_GROWTH"]):
            e *= 1 + gr
            divs += e * payout
        future = e * pe_term + divs
        tae = (future / price) ** 0.1 - 1 if future > 0 else None
        price_15 = future / (1 + cfg["TARGET_RETURN"]) ** 10
        row["PER salida"] = _rnd(pe_term, 1)

    # ----- PASO 5: precio, margen de seguridad y señal -----
    ms_now = (1 - price / iv) if (_ok(iv) and iv > 0 and _ok(price)) else None
    oe_yield = (oe_ps / price) if (_ok(oe_ps) and _ok(price) and price > 0) else None
    f4_items = [
        (bracket(ms_now, [(0.40, 14), (0.30, 12), (0.25, 10), (0.15, 7), (0.0, 4), (-0.20, 1)]) or 0, 14),
        (bracket(tae, [(0.15, 8), (0.12, 6), (0.10, 4), (0.08, 2)]) or 0, 8),
        (bracket((oe_yield / ust10) if (_ok(oe_yield) and ust10) else None, [(2.0, 8), (1.5, 6), (1.0, 3)]) or 0, 8),
    ]
    F4 = float(sum(pts for pts, _ in f4_items))

    quality = F1 + F2 + F3
    score = quality + F4
    if quality >= cfg["QUALITY_WONDERFUL"]:
        clase, ms_req = "Maravillosa", cfg["MS_WONDERFUL"]
    elif quality >= cfg["QUALITY_GOOD"]:
        clase, ms_req = "Buena", cfg["MS_GOOD"]
    else:
        clase, ms_req = "Mediocre", None

    hard = []
    if n < 5:
        hard.append("Historial < 5 años")
    if n >= 5 and frac_profit < 0.7:
        hard.append("Pérdidas en > 30% de los años")
    if not (_ok(ni3) and ni3 > 0):
        hard.append("Beneficio medio 3 años ≤ 0")
    if min(cov1, cov2, cov3) < 0.6:
        alerts.append("Datos insuficientes en algún factor")

    buy_price = iv * (1 - ms_req) if (_ok(iv) and ms_req is not None) else None
    dist = (price / buy_price - 1) if (_ok(buy_price) and buy_price > 0 and _ok(price)) else None

    if hard or clase == "Mediocre":
        signal = "DESCARTAR"
    elif not _ok(price):
        signal = "REVISAR"
        alerts.append("Sin precio")
    elif not _ok(iv):
        signal = "REVISAR"
        alerts.append("Owner earnings ≤ 0: no valorable por flujos")
    elif price <= buy_price and (tae is None or tae >= cfg["MIN_TAE_BUY"]):
        signal = "COMPRAR"
    elif price <= buy_price * (1 + cfg["WATCH_BAND"]):
        signal = "VIGILAR"
    elif price >= iv * cfg["OVERVALUED"]:
        signal = "SOBREVALORADA"
    else:
        signal = "CARO"
    if signal == "COMPRAR" and ((_ok(ms_now) and ms_now > cfg["ANOMALY_MS"]) or
                                (_ok(oe_yield) and oe_yield > cfg["ANOMALY_OE_YIELD"])):
        signal = "REVISAR"
        alerts.append("Valoración anómala: verificar datos en el 10-K antes de comprar")

    # ----- Analistas (informativo, 0 puntos) -----
    tm = mkt.get("target_mean")
    up_an = (tm / price - 1) if (_ok(tm) and _ok(price) and price > 0) else None
    rm = mkt.get("rec_mean")
    votes = mkt.get("votes")
    ets = mkt.get("earnings_ts")
    next_er = None
    if _ok(ets):
        d = dt.datetime.fromtimestamp(float(ets), dt.timezone.utc).date()
        if d >= TODAY - dt.timedelta(days=2):
            next_er = d.isoformat()
            if (d - TODAY).days <= 7:
                alerts.append("Resultados en ≤ 7 días: el VI puede cambiar")

    row.update({
        "Puntuación": _rnd(score, 1), "Calidad": _rnd(quality, 1), "Clase": clase,
        "F1 Comprensible": _rnd(F1, 1), "F2 Foso": _rnd(F2, 1), "F3 Dirección": _rnd(F3, 1), "F4 Precio": _rnd(F4, 1),
        "Señal": signal,
        "Valor intrínseco": _rnd(iv), "Precio de compra": _rnd(buy_price),
        "MS exigido %": _pct(ms_req, 0) if ms_req is not None else None,
        "MS actual %": _pct(ms_now), "Dist. a precio compra %": _pct(dist),
        "Upside VI %": _pct((iv / price - 1) if (_ok(iv) and _ok(price) and price > 0) else None),
        "TAE 10a %": _pct(tae), "Precio para 15%": _rnd(price_15),
        "OE por acción": _rnd(oe_ps), "OE yield %": _pct(oe_yield), "Bono 10a %": _pct(ust10, 2),
        "Tasa descuento %": _pct(r_disc, 2),
        "Consenso": mkt.get("rec_key"), "Rec. media": _rnd(rm, 2), "Nº analistas": mkt.get("n_analysts"),
        "Votos C/M/V": f"{votes[0]}/{votes[1]}/{votes[2]}" if votes else None,
        "Obj. medio": _rnd(tm), "Upside analistas %": _pct(up_an),
        "Obj. bajo": _rnd(mkt.get("target_low")), "Obj. alto": _rnd(mkt.get("target_high")),
        "Próx. resultados": next_er,
        "ROE 10a %": None if neg_equity else _pct(roe_avg), "ROIC 10a %": _pct(roic_avg), "ROA 10a %": _pct(roa_avg, 2),
        "Margen bruto %": _pct(gm_avg), "Margen neto %": _pct(nm_avg),
        "Tendencia ROIC (pp)": _rnd(roic_trend, 1), "Tendencia margen bruto (pp)": _rnd(gm_trend, 1),
        "Crec. BPA %": _pct(eps_cagr), "g usada %": _pct(g), "Capex/BN %": _pct(capex_ni),
        "Conversión FCF %": _pct(fcf_conv), "Test dólar retenido": _rnd(retained_ratio),
        "CAGR acciones %": _pct(sh_cagr, 2), "Deuda/BN (años)": _rnd(debt_years, 1),
        "SBC/FCF %": _pct(sbc_fcf), "Devengos %": _pct(accr), "Años con beneficio": yrs_profit,
        "Caídas ingresos >10%": drops, "Círculo (prior)": circle, "Cíclica": "Sí" if cyclical else "No",
        "Filtros duros": "; ".join(hard), "Alertas": "; ".join(dict.fromkeys(alerts)),
    })
    return row


# =====================================================================
# 5. REVISIONES CUALITATIVAS (memoria del agente)
# =====================================================================
REVIEW_COLS = ["ticker", "fecha", "circulo", "foso", "direccion", "ajuste", "motivo"]


def load_reviews(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame(columns=REVIEW_COLS)
    rv = pd.read_csv(path, dtype=str, sep=None, engine="python").fillna("")
    rv.columns = [c.strip().lower() for c in rv.columns]
    for c in REVIEW_COLS:
        if c not in rv.columns:
            rv[c] = ""
    rv["ticker"] = rv["ticker"].str.strip().str.upper()
    rv["fecha"] = pd.to_datetime(rv["fecha"], errors="coerce")
    rv = rv.dropna(subset=["fecha"]).sort_values("fecha").drop_duplicates("ticker", keep="last")
    return rv[REVIEW_COLS]


def weight_for(score):
    return "8–10 %" if score >= 80 else ("5–7 %" if score >= 70 else "3–5 %")


def apply_reviews(df, reviews, cfg, data_date):
    """Añade Foso, Ajuste, Punt. final y Revisado; aplica los topes de señal del juicio cualitativo."""
    df = df.copy()
    for c in ["Foso", "Ajuste", "Revisado", "Revisión caducada"]:
        df[c] = None
    df["Punt. final"] = df["Puntuación"]
    if len(reviews):
        today = pd.Timestamp(data_date)
        rv = reviews.set_index("ticker")
        for i, r in df.iterrows():
            t = str(r["Ticker"]).upper()
            if t not in rv.index:
                continue
            x = rv.loc[t]
            try:
                adj = float(str(x["ajuste"]).replace(",", ".").replace("+", "") or 0)
            except ValueError:
                adj = 0.0
            adj = float(np.clip(adj, -10, 5))
            age = (today - x["fecha"]).days
            df.at[i, "Foso"] = x["foso"] or None
            df.at[i, "Ajuste"] = adj
            df.at[i, "Revisado"] = x["fecha"].date().isoformat()
            df.at[i, "Revisión caducada"] = "Sí" if age > cfg["REVIEW_MAX_DAYS"] else "No"
            if _ok(r["Puntuación"]):
                df.at[i, "Punt. final"] = round(float(r["Puntuación"]) + adj, 1)
            circ, dire, foso = x["circulo"].strip().lower(), x["direccion"].strip().lower(), x["foso"].strip().lower()
            if circ.startswith("fuera") or dire.startswith("no fiable"):
                df.at[i, "Señal"] = "DESCARTAR"
            elif foso.startswith("ninguno") and r["Señal"] == "COMPRAR":
                df.at[i, "Señal"] = "VIGILAR"
    df["Peso sugerido"] = [weight_for(s) if (sig == "COMPRAR" and _ok(s)) else ""
                           for s, sig in zip(df["Punt. final"], df["Señal"])]
    rm, up = df["Rec. media"], df["Upside analistas %"]
    sent = []
    for sig, a, b in zip(df["Señal"], rm, up):
        s = "Neutral"
        if sig in ("COMPRAR", "VIGILAR") and ((_ok(a) and a >= 2.6) or (_ok(b) and b < 5)):
            s = "Miedo (contrario)"
        elif sig in ("CARO", "SOBREVALORADA") and _ok(a) and a <= 2.0 and _ok(b) and b > 15:
            s = "Euforia"
        sent.append(s)
    df["Sentimiento"] = sent
    return df


# =====================================================================
# 6. ORQUESTACIÓN Y SALIDAS
# =====================================================================
MAIN_COLS = [
    "Rank", "Ticker", "Empresa", "Sector", "Precio", "Punt. final", "Puntuación", "Ajuste", "Calidad", "Clase",
    "F1 Comprensible", "F2 Foso", "F3 Dirección", "F4 Precio", "Foso", "Señal", "Peso sugerido",
    "Valor intrínseco", "Precio de compra", "MS exigido %", "MS actual %", "Dist. a precio compra %",
    "Upside VI %", "TAE 10a %", "Precio para 15%", "OE yield %", "Bono 10a %", "Tasa descuento %",
    "Consenso", "Rec. media", "Nº analistas", "Votos C/M/V", "Obj. medio", "Upside analistas %",
    "Obj. bajo", "Obj. alto", "Sentimiento", "Próx. resultados", "Revisado", "Revisión caducada",
    "ROE 10a %", "ROIC 10a %", "ROA 10a %", "Margen bruto %", "Margen neto %",
    "Tendencia ROIC (pp)", "Tendencia margen bruto (pp)", "Crec. BPA %", "g usada %",
    "Capex/BN %", "Conversión FCF %", "Test dólar retenido", "CAGR acciones %", "Deuda/BN (años)",
    "SBC/FCF %", "Devengos %", "OE por acción", "PER salida", "Años historia", "Años con beneficio",
    "Caídas ingresos >10%", "Cíclica", "Círculo (prior)", "Modo", "Subsector", "Filtros duros", "Alertas", "Fuente",
]
AGENT_COLS = [
    "Rank", "Ticker", "Empresa", "Sector", "Precio", "Punt. final", "Puntuación", "Ajuste", "Calidad", "Clase",
    "F1 Comprensible", "F2 Foso", "F3 Dirección", "F4 Precio", "Foso", "Señal", "Peso sugerido",
    "Valor intrínseco", "Precio de compra", "MS actual %", "Dist. a precio compra %", "Upside VI %", "TAE 10a %",
    "OE yield %", "Obj. medio", "Upside analistas %", "Votos C/M/V", "Rec. media", "Sentimiento",
    "Próx. resultados", "Revisado", "Revisión caducada", "Tendencia ROIC (pp)", "Tendencia margen bruto (pp)",
    "Cíclica", "Alertas",
]
SIGNAL_ORDER = {"COMPRAR": 0, "VIGILAR": 1, "CARO": 2, "SOBREVALORADA": 3, "REVISAR": 4, "DESCARTAR": 5}


def rank_rows(df, data_date):
    df = df.copy()
    df["_d"] = (df["Señal"] == "DESCARTAR").astype(int)
    df = df.sort_values(["_d", "Punt. final"], ascending=[True, False]).drop(columns="_d").reset_index(drop=True)
    df.insert(0, "Rank", range(1, len(df) + 1))
    df.insert(0, "Fecha datos", data_date)
    cols = ["Fecha datos"] + [c for c in MAIN_COLS if c in df.columns]
    return df[cols]


def market_summary(df, ust10, data_date, n_universe, cfg, reviews_hash):
    valid = df[df["Valor intrínseco"].notna() & df["Precio"].notna()]
    ratio = valid["Precio"] / valid["Valor intrínseco"]
    counts = df["Señal"].value_counts().to_dict()
    return {
        "Fecha datos": data_date, "Generado (UTC)": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "Bono 10a %": _pct(ust10, 2), "Tasa descuento %": _pct(discount_rate(cfg, ust10), 2),
        "Empresas universo": n_universe, "Empresas analizadas": len(df), "Con valor intrínseco": len(valid),
        "Mediana Precio/VI": _rnd(ratio.median(), 2), "% cotizando bajo VI": _pct((ratio < 1).mean()),
        **{f"Nº {k}": counts.get(k, 0) for k in SIGNAL_ORDER},
        "Revisiones aplicadas": int(df["Revisado"].notna().sum()),
        "Revisiones caducadas": int((df["Revisión caducada"] == "Sí").sum()),
        "Hash revisiones": reviews_hash,
    }


def changes_vs_previous(df, hist_dir, data_date):
    prev = sorted(f for f in os.listdir(hist_dir) if f.endswith(".csv") and f[:10] < data_date) if os.path.isdir(hist_dir) else []
    if not prev:
        return pd.DataFrame()
    old = pd.read_csv(os.path.join(hist_dir, prev[-1]))
    m = df[["Ticker", "Señal", "Punt. final", "Precio"]].merge(
        old[["Ticker", "Señal", "Punt. final"]], on="Ticker", how="outer", suffixes=("", " anterior"))
    ch = m[m["Señal"] != m["Señal anterior"]].copy()
    ch["Fecha anterior"] = prev[-1][:10]
    return ch.sort_values("Señal", key=lambda s: s.map(SIGNAL_ORDER).fillna(9))


def export(df, summary, changes, cfg, data_date):
    out = cfg["OUTPUT_DIR"]
    hist = os.path.join(out, "historico")
    os.makedirs(hist, exist_ok=True)
    df.to_csv(os.path.join(out, "screener_buffett.csv"), index=False, encoding="utf-8-sig")
    agent = df[df["Señal"] != "DESCARTAR"][["Fecha datos"] + AGENT_COLS]
    agent.to_csv(os.path.join(out, "screener_buffett_agente.csv"), index=False, encoding="utf-8-sig")
    agent.to_csv(os.path.join(hist, f"{data_date}.csv"), index=False, encoding="utf-8-sig")
    changes.to_csv(os.path.join(out, "cambios.csv"), index=False, encoding="utf-8-sig")
    with open(os.path.join(out, "mercado.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=1)
    if cfg["WRITE_XLSX"]:
        xlsx_path = os.path.join(out, f"screener_buffett_{data_date}.xlsx")
        top = pd.concat([df[df["Señal"].isin(["COMPRAR", "VIGILAR"])], df.head(25)]).drop_duplicates("Ticker")
        params = pd.DataFrame([(k, str(v)) for k, v in cfg.items() if k != "SEC_USER_AGENT"], columns=["Parámetro", "Valor"])
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as xw:
            df.to_excel(xw, sheet_name="Screener", index=False)
            top.to_excel(xw, sheet_name="Top", index=False)
            (changes if len(changes) else pd.DataFrame({"Info": ["Sin día anterior o sin cambios de señal"]})).to_excel(
                xw, sheet_name="Cambios", index=False)
            pd.DataFrame(list(summary.items()), columns=["Métrica", "Valor"]).to_excel(xw, sheet_name="Mercado", index=False)
            params.to_excel(xw, sheet_name="Parámetros", index=False)
            _format_xlsx(xw)
    with open(os.path.join(out, "mensaje.txt"), "w", encoding="utf-8") as f:
        f.write(f"buffett {data_date}: {summary['Empresas analizadas']} empresas, "
                f"{summary['Nº COMPRAR']} COMPRAR, {summary['Nº VIGILAR']} VIGILAR")


def _format_xlsx(xw):
    from openpyxl.styles import Font, PatternFill
    fills = {"COMPRAR": "C6EFCE", "VIGILAR": "FFEB9C", "CARO": "F2F2F2", "SOBREVALORADA": "FFC7CE",
             "REVISAR": "DDEBF7", "DESCARTAR": "E7E6E6"}
    for name, ws in xw.sheets.items():
        ws.freeze_panes = "D2" if name in ("Screener", "Top") else "A2"
        for c in ws[1]:
            c.font = Font(bold=True)
        headers = [c.value for c in ws[1]]
        for col_idx, h in enumerate(headers, 1):
            ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = max(10, min(40, len(str(h)) + 2))
        if "Señal" in headers:
            si = headers.index("Señal") + 1
            for r in range(2, ws.max_row + 1):
                v = ws.cell(r, si).value
                if v in fills:
                    ws.cell(r, si).fill = PatternFill("solid", fgColor=fills[v])


def _file_hash(path):
    if not path or not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return hashlib.sha1(f.read()).hexdigest()[:12]


def _gh_output(work):
    """Dice al workflow si hubo trabajo (para guardar la caché solo entonces)."""
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as f:
            f.write(f"work={1 if work else 0}\n")


def prewarm(cfg):
    """Precarga de la tarde: lista, SEC, histórico mensual y analistas. No calcula ni publica nada."""
    t0 = time.time()
    uni = get_universe(cfg)
    if cfg.get("LIMIT"):
        uni = uni.head(cfg["LIMIT"])
    tickers = uni["yf"].tolist()
    log("Precarga 1/3: SEC")
    sec = SecFrames(cfg)
    build_fundamentals(uni["cik"].tolist(), list(range(TODAY.year - 11, TODAY.year + 1)), sec.get)
    log(f"  peticiones SEC nuevas: {sec.calls}")
    log("Precarga 2/3: histórico mensual")
    get_prices(tickers, cfg)
    log("Precarga 3/3: analistas")
    an = get_analysts(tickers, cfg)
    ok = sum(1 for t in tickers if an.get(t, {}).get("market_cap"))
    log(f"Precarga lista en {(time.time() - t0) / 60:.1f} min: analistas {ok}/{len(tickers)}")
    _gh_output(True)


def run(cfg=CONFIG):
    ua = cfg["SEC_USER_AGENT"]
    if "tu_email@ejemplo.com" in ua or "@" not in ua:
        raise ValueError("Falta SEC_USER_AGENT (nombre y email). En GitHub: secreto del repositorio; en Colab: CONFIG.")
    os.makedirs(cfg["OUTPUT_DIR"], exist_ok=True)
    if cfg["PREWARM"]:
        prewarm(cfg)
        return None
    reviews_path = cfg["REVIEWS_FILE"] or os.path.join(cfg["OUTPUT_DIR"], "revisiones.csv")
    t0 = time.time()

    session = get_session()
    rhash = _file_hash(reviews_path)
    prev_file = os.path.join(cfg["OUTPUT_DIR"], "mercado.json")
    if os.path.exists(prev_file) and not cfg["FORCE"]:
        with open(prev_file, encoding="utf-8") as f:
            prev = json.load(f)
        if prev.get("Fecha datos") == session and prev.get("Hash revisiones", "") == rhash:
            log(f"La sesión {session} ya está calculada y las revisiones no han cambiado. Nada que hacer.")
            _gh_output(False)
            return None
    log(f"Sesión de referencia: {session}")

    log("1/6 Universo S&P 500")
    uni = get_universe(cfg)
    if cfg.get("LIMIT"):
        uni = uni.head(cfg["LIMIT"])
    tickers = uni["yf"].tolist()

    log("2/6 Fundamentales SEC (10–12 años)")
    years = list(range(TODAY.year - 11, TODAY.year + 1))
    sec = SecFrames(cfg)
    panels = build_fundamentals(uni["cik"].tolist(), years, sec.get)
    log(f"  peticiones SEC nuevas: {sec.calls}")

    log("3/6 Precios, histórico y bono 10 años")
    last, last_date, monthly, splits = get_prices(tickers, cfg)
    ust10 = get_ust10(cfg)
    data_date = last_date.isoformat()

    log("4/6 Analistas y capitalización (Yahoo)")
    an = get_analysts(tickers, cfg)

    log("5/6 Puntuación Buffett")
    rows = []
    for _, u in uni.iterrows():
        tk = u["yf"]
        panel = panels.get(int(u["cik"]))
        fuente = "SEC"
        if panel is None or panel.get("net_income", pd.Series(dtype=float)).notna().sum() < 5:
            try:
                alt = yf_statements(tk)
                if alt is not None and (panel is None or alt["net_income"].notna().sum() > panel.get("net_income", pd.Series(dtype=float)).notna().sum()):
                    panel, fuente = alt, "Yahoo (historial corto)"
            except Exception:  # noqa: BLE001
                pass
        mkt = dict(an.get(tk, {}))
        mkt["price"] = float(last.get(tk)) if _ok(last.get(tk)) else None
        mpx = monthly[tk] if tk in getattr(monthly, "columns", []) else None
        sby = {}
        if isinstance(splits, pd.DataFrame) and tk in splits.columns:
            sp = splits[tk].dropna()
            for d, v in sp[sp > 0].items():
                sby[d.year] = sby.get(d.year, 1.0) * float(v)
        meta = {"ticker": u["ticker"], "empresa": u["empresa"], "sector": u["sector"],
                "subsector": u["subsector"], "fuente": fuente}
        try:
            rows.append(compute_company(meta, panel, mkt, ust10, mpx, sby, cfg))
        except Exception as e:  # noqa: BLE001
            rows.append({"Ticker": u["ticker"], "Empresa": u["empresa"], "Sector": u["sector"],
                         "Subsector": u["subsector"], "Señal": "REVISAR", "Puntuación": None,
                         "Alertas": f"Error de cálculo: {e}"})

    log("6/6 Revisiones cualitativas y salidas")
    df = apply_reviews(pd.DataFrame(rows), load_reviews(reviews_path), cfg, data_date)
    df = rank_rows(df, data_date)
    summary = market_summary(df, ust10, data_date, len(uni), cfg, rhash)
    changes = changes_vs_previous(df, os.path.join(cfg["OUTPUT_DIR"], "historico"), data_date)
    export(df, summary, changes, cfg, data_date)
    log(f"Listo en {(time.time() - t0) / 60:.1f} min")
    _gh_output(True)
    for k, v in summary.items():
        print(f"  {k}: {v}")
    return df, summary, changes


if __name__ == "__main__" and "ipykernel" not in __import__("sys").modules:
    run()

# %%
# from google.colab import drive; drive.mount("/content/drive")   # ← descomenta en Colab
# out = run()
# if out: out[0][out[0]["Señal"].isin(["COMPRAR", "VIGILAR"])].head(30)
