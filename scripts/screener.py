#!/usr/bin/env python3
"""
HalalStocks Automated Screener
Fetches financial data via yfinance and applies AAOIFI Shariah screening criteria.
Run quarterly via GitHub Actions to keep stocks.json up to date.
"""

import json
import os
import sys
import time
from datetime import datetime

try:
    import yfinance as yf
except ImportError:
    print("ERROR: yfinance not installed. Run: pip install -r requirements.txt")
    sys.exit(1)

# ─── TICKER LISTS BY EXCHANGE ──────────────────────────────────────────────

TICKERS_US = [
    # Tech
    "AAPL","MSFT","GOOGL","NVDA","TSLA","META","AMZN","AMD","ADBE","CRM",
    "INTC","IBM","ORCL","NET","SNOW","ABNB","SHOP","PLTR","UBER","DELL",
    # Healthcare
    "JNJ","UNH","PFE","ABT","TMO","DHR","ISRG","SYK","MDT","BSX",
    "GEHC","HCA","CVS",
    # Finance
    "JPM","BAC","WFC","GS","MS","C","AXP","BLK","SCHW","COF","V","MA","BRK-B",
    # Consumer
    "KO","PEP","MCD","SBUX","CMG","YUM","KHC","WMT","COST","TGT","HD","LOW",
    "NKE","EL",
    # Industrial
    "GE","HON","MMM","CAT","DE","ETN","ITW","BA","RTX","LMT","NOC",
    # Energy/Materials
    "XOM","CVX","SLB","HAL","LIN","APD","NEM","FCX",
    # Telecom/Media
    "T","VZ","TMUS","CMCSA","DIS","NFLX",
    # Tobacco
    "MO","PM",
]

TICKERS_DE = [
    "SAP.DE","SIE.DE","ALV.DE","MRK.DE","BAS.DE","DTE.DE","BMW.DE",
    "MBG.DE","ADS.DE","BAYN.DE","HEN3.DE","IFX.DE","LHA.DE","MUV2.DE",
    "BEI.DE","SHL.DE","VOW3.DE","PAH3.DE","DBK.DE","CBK.DE","ZAL.DE",
    "FRE.DE","AIR.PA","RWE.DE","ENR.DE","MTX.DE",
]

TICKERS_SA = [
    "2222.SR","1120.SR","7010.SR","2010.SR","1150.SR","1140.SR",
    "2280.SR","4190.SR","7020.SR","2350.SR","4004.SR","2380.SR",
    "1010.SR","1030.SR",
]

TICKERS_AE = [
    "EMAAR.AE","FAB.AE","DIB.AE","ADCB.AE","EAND.AE","DU.AE",
    "ALDAR.AE","DEWA.AE",
]

TICKERS_MY = [
    "1155.KL","1023.KL","5681.KL","5347.KL","1295.KL","5183.KL",
    "5099.KL","7084.KL","3816.KL","7113.KL","5285.KL","5819.KL",
]

TICKERS_TR = [
    "GARAN.IS","AKBNK.IS","TCELL.IS","BIMAS.IS","KCHOL.IS",
    "THYAO.IS","EREGL.IS","TUPRS.IS","ARCLK.IS","SISE.IS",
    "TTKOM.IS","FROTO.IS","ASELS.IS",
]

# ─── EXCHANGE & COUNTRY MAPPING ────────────────────────────────────────────

TICKER_META = {
    # US
    **{t: {"ex": "NASDAQ" if t in ["AAPL","MSFT","GOOGL","NVDA","TSLA","META","AMZN","AMD","ADBE","INTC","PEP","COST","NFLX","TMUS","CMCSA","PLTR","SNOW","ABNB","SHOP","UBER","KHC","HON"] else "NYSE", "ctry": "US", "flag": "🇺🇸"} for t in TICKERS_US},
    "BRK-B": {"ex": "NYSE", "ctry": "US", "flag": "🇺🇸"},
    # Germany
    **{t: {"ex": "XETRA", "ctry": "DE", "flag": "🇩🇪"} for t in TICKERS_DE},
    "AIR.PA": {"ex": "XETRA", "ctry": "DE", "flag": "🇩🇪"},
    # Saudi
    **{t: {"ex": "TASI", "ctry": "SA", "flag": "🇸🇦"} for t in TICKERS_SA},
    # UAE
    **{t: {"ex": "DFM", "ctry": "AE", "flag": "🇦🇪"} for t in TICKERS_AE},
    # Malaysia
    **{t: {"ex": "BURSA", "ctry": "MY", "flag": "🇲🇾"} for t in TICKERS_MY},
    # Turkey
    **{t: {"ex": "BIST", "ctry": "TR", "flag": "🇹🇷"} for t in TICKERS_TR},
}

# ─── SHARIAH CLASSIFICATION ────────────────────────────────────────────────

# Industries whose PRIMARY business is haram
HARAM_INDUSTRIES = {
    "banks—regional", "banks—diversified", "banks", "bank",
    "investment banking & brokerage", "investment banking",
    "credit services", "consumer finance",
    "insurance—life", "insurance—property & casualty", "insurance—diversified",
    "insurance", "reinsurance",
    "casinos & gaming", "gambling", "lottery",
    "tobacco", "cigarettes",
    "beverages—wineries & distilleries", "beverages—brewers",
    "brewers", "distillers",
    "adult entertainment",
}

# Keywords in company name / industry that signal haram
HARAM_KEYWORDS = [
    "bank", "bancorp", "banque", "financial corp", "financial group",
    "insurance", "assurance", "reinsurance",
    "casino", "gambling", "lottery",
    "tobacco", "cigarette", "cigar",
    "brewery", "brewing", "distillery", "winery", "wine", "spirits", "alcohol", "beer",
]

# Industries that are doubtful (need ratio checks)
DOUBTFUL_INDUSTRIES = {
    "oil & gas e&p", "oil & gas integrated", "oil & gas refining & marketing",
    "oil & gas midstream", "oil & gas equipment & services",
    "coal", "thermal coal",
    "aerospace & defense", "defense",
    "specialty retail", "grocery stores", "discount stores", "department stores",
    "restaurants", "fast food",
    "media—diversified", "broadcasting", "entertainment",
    "airlines",
    "hotels & motels", "resorts & casinos",
    "specialty finance", "capital markets",
}

# Estimated haram income % by industry
HARAM_INCOME_ESTIMATE = {
    "grocery stores": 4.5,
    "discount stores": 3.8,
    "department stores": 3.2,
    "specialty retail": 1.5,
    "restaurants": 2.5,
    "airlines": 3.5,
    "hotels & motels": 3.0,
    "media—diversified": 4.0,
    "broadcasting": 3.0,
    "beverages—non-alcoholic": 2.0,
    "personal products": 1.0,
    "packaged foods": 1.0,
    "aerospace & defense": 40.0,
    "tobacco": 100.0,
    "casinos & gaming": 100.0,
    "beverages—wineries & distilleries": 100.0,
    "beverages—brewers": 100.0,
}

# ─── SCREENER LOGIC ────────────────────────────────────────────────────────

def safe_ratio(num, denom, pct=True):
    """Safely divide and optionally convert to percentage."""
    try:
        if not num or not denom or denom == 0:
            return 0.0
        ratio = num / denom
        return round(ratio * 100 if pct else ratio, 1)
    except Exception:
        return 0.0


def classify_business(info):
    """Return (status_char, b_criterion, reason_hint) based on industry/sector."""
    industry = (info.get("industry") or "").lower()
    sector = (info.get("sector") or "").lower()
    name = (info.get("longName") or info.get("shortName") or "").lower()

    # Check haram keywords in name
    for kw in HARAM_KEYWORDS:
        if kw in name and kw not in ["financial corp", "financial group"]:
            pass  # name alone not always deterministic

    # Explicit haram industries
    for hi in HARAM_INDUSTRIES:
        if hi in industry:
            return "haram", "fail"

    # Defense check
    if "aerospace & defense" in industry or "defense" in industry:
        return "haram_candidate", "fail"

    # Doubtful industries
    for di in DOUBTFUL_INDUSTRIES:
        if di in industry:
            return "doubtful", "warn"

    return "halal", "pass"


def get_debt_ratio(info):
    """Total debt / Market cap as %."""
    debt = info.get("totalDebt") or 0
    mc = info.get("marketCap") or 0
    if mc == 0:
        return None
    return round((debt / mc) * 100, 1)


def get_interest_income_ratio(info, balance_sheet=None):
    """Interest income as % of total revenue."""
    revenue = info.get("totalRevenue") or 0
    if revenue == 0:
        return 0.0
    interest_income = info.get("interestIncome") or 0
    if interest_income < 0:
        interest_income = abs(interest_income)
    return round((interest_income / revenue) * 100, 1)


def get_haram_income_estimate(info):
    """Estimate haram income % from industry classification."""
    industry = (info.get("industry") or "").lower()
    for ind, pct in HARAM_INCOME_ESTIMATE.items():
        if ind in industry:
            return pct
    return 0.0


def compute_criteria(debt_ratio, interest_ratio, haram_pct, biz_status):
    """Compute pass/warn/fail for each criterion."""
    # Primary business
    b = biz_status  # already pass/warn/fail

    # Debt ratio
    if debt_ratio is None:
        d = "fail"  # banks have N/A but are already haram
    elif debt_ratio < 20:
        d = "pass"
    elif debt_ratio < 30:
        d = "warn"
    else:
        d = "fail"

    # Interest income
    if interest_ratio < 2.5:
        i = "pass"
    elif interest_ratio < 5:
        i = "warn"
    else:
        i = "fail"

    # Haram income
    if haram_pct < 2.5:
        h = "pass"
    elif haram_pct < 5:
        h = "warn"
    else:
        h = "fail"

    return {"b": b, "d": d, "i": i, "h": h}


def compute_status(cr, biz_type):
    """Compute overall halal/doubtful/haram status."""
    if biz_type in ("haram", "haram_candidate") or cr["b"] == "fail":
        return "r"
    if cr["h"] == "fail":
        return "r"
    if cr["i"] == "fail":
        return "r"
    if cr["d"] == "fail" or cr["b"] == "warn" or cr["h"] == "warn" or cr["i"] == "warn":
        return "d"
    return "h"


SECTOR_MAP_RU = {
    "Technology": "Технологии",
    "Healthcare": "Здравоохранение",
    "Financial Services": "Финансовые услуги",
    "Consumer Cyclical": "Потребительский сектор",
    "Consumer Defensive": "Потребительские товары",
    "Industrials": "Промышленность",
    "Energy": "Энергетика",
    "Basic Materials": "Материалы",
    "Communication Services": "Телекоммуникации",
    "Utilities": "Коммунальные услуги",
    "Real Estate": "Недвижимость",
}

SECTOR_MAP_TJ = {
    "Technology": "Технологияҳо",
    "Healthcare": "Тандурустӣ",
    "Financial Services": "Хидматҳои молиявӣ",
    "Consumer Cyclical": "Бахши истеъмолӣ",
    "Consumer Defensive": "Молҳои истеъмолӣ",
    "Industrials": "Саноат",
    "Energy": "Энергетика",
    "Basic Materials": "Маводҳои асосӣ",
    "Communication Services": "Телекоммуникатсия",
    "Utilities": "Хизматрасониҳои коммуналӣ",
    "Real Estate": "Амволи ғайриманқул",
}

# ─── REASON GENERATOR ──────────────────────────────────────────────────────

def build_reason(info, sts, cr, debt_ratio, interest_ratio, haram_pct):
    """Build Russian reason text."""
    name = info.get("longName") or info.get("shortName") or "Компания"
    industry = info.get("industry") or "N/A"
    sector = info.get("sector") or "N/A"
    sector_ru = SECTOR_MAP_RU.get(sector, sector)

    if sts == "r":
        if "bank" in industry.lower() or "insurance" in industry.lower():
            return f"{name} — конвенциональный банк/страховщик. Основная деятельность строится на риба (процентах) или gharar (неопределённости), что прямо запрещено шариатом по всем мазхабам."
        if "tobacco" in industry.lower():
            return f"{name} производит табачные изделия, что запрещено шариатом как причинение вреда здоровью. Харамный доход составляет 100% выручки."
        if "defense" in industry.lower() or "aerospace" in industry.lower():
            return f"{name} — производитель военного вооружения. Основная выручка от контрактов на производство систем вооружения, что считается недопустимым большинством шариатских учёных."
        if haram_pct > 50:
            return f"{name}: основная часть выручки ({haram_pct:.0f}%) поступает из харамных источников. Несоответствие критерию AAOIFI по харамному доходу."
        return f"{name} не соответствует критериям AAOIFI. Основная деятельность или структура доходов несовместимы с нормами шариата."

    if sts == "d":
        reasons = []
        if cr["d"] == "fail":
            reasons.append(f"долговая нагрузка {debt_ratio:.0f}% превышает допустимый порог AAOIFI в 30%")
        if cr["b"] == "warn":
            reasons.append(f"основной бизнес в секторе \"{industry}\" требует дополнительного изучения")
        if cr["h"] == "warn":
            reasons.append(f"оценочный харамный доход ({haram_pct:.1f}%) близок к пороговому значению 5%")
        if cr["i"] == "fail":
            reasons.append(f"процентный доход ({interest_ratio:.1f}%) превышает допустимый порог AAOIFI в 5%")
        if cr["i"] == "warn":
            reasons.append(f"процентный доход ({interest_ratio:.1f}%) приближается к порогу 5%")
        if not reasons:
            reasons.append("некоторые аспекты деятельности требуют проверки")
        reason_str = "; ".join(reasons[:2])
        return f"{name} ({sector_ru}): {reason_str}. Рекомендуется консультация с шариатским учёным."

    # halal
    debt_str = f"{debt_ratio:.0f}%" if debt_ratio is not None else "N/A"
    return (f"{name} ведёт деятельность в секторе {sector_ru}. "
            f"Долговая нагрузка ({debt_str}) соответствует нормам, харамный доход отсутствует. "
            f"Полностью соответствует критериям AAOIFI.")


def build_reason_tj(info, sts, cr, debt_ratio, haram_pct):
    """Build Tajik reason (condensed)."""
    name = info.get("longName") or info.get("shortName") or "Ширкат"
    if sts == "r":
        if "bank" in (info.get("industry") or "").lower():
            return f"{name} — бонки анъанавӣ. Фаъолияти асосӣ риба аст, ки шариат манъ мекунад."
        return f"{name} ба меъёрҳои AAOIFI мувофиқат намекунад. Фаъолият ё даромад бо шариат созгор нест."
    if sts == "d":
        if cr["d"] == "fail":
            return f"{name}: бори қарзӣ ({debt_ratio:.0f}%) аз ҳадди 30% AAOIFI зиёд аст. Машварат бо олими шариат зарур аст."
        return f"{name}: баъзе ҷанбаҳои фаъолият санҷиши иловагиро талаб мекунанд. Машварат тавсия дода мешавад."
    debt_str = f"{debt_ratio:.0f}%" if debt_ratio is not None else "N/A"
    return f"{name} ба меъёрҳои AAOIFI мувофиқат мекунад. Бори қарзӣ {debt_str}, даромади ҳаром нест."


# ─── MAIN SCREENER ─────────────────────────────────────────────────────────

def screen_ticker(ticker):
    """Screen a single ticker. Returns dict or None on failure."""
    meta = TICKER_META.get(ticker, {"ex": "NYSE", "ctry": "US", "flag": "🇺🇸"})

    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        if not info or not info.get("sector"):
            print(f"  [SKIP] {ticker}: no info available")
            return None

        # Financial ratios
        debt_ratio = get_debt_ratio(info)
        interest_ratio = get_interest_income_ratio(info)
        haram_pct = get_haram_income_estimate(info)

        # Business classification
        biz_type, b_crit = classify_business(info)

        # Criteria
        cr = compute_criteria(debt_ratio, interest_ratio, haram_pct, b_crit)

        # Status
        sts = compute_status(cr, biz_type)

        # Names & sectors
        name = info.get("longName") or info.get("shortName") or ticker
        sector = info.get("sector") or "Other"
        sr = SECTOR_MAP_RU.get(sector, sector)
        st = SECTOR_MAP_TJ.get(sector, sector)

        # Reasons
        rr = build_reason(info, sts, cr, debt_ratio, interest_ratio, haram_pct)
        rt = build_reason_tj(info, sts, cr, debt_ratio, haram_pct)

        return {
            "tk": ticker,
            "ex": meta["ex"],
            "ctry": meta["ctry"],
            "flag": meta["flag"],
            "nr": name,
            "nt": name,
            "sr": sr,
            "st": st,
            "sts": sts,
            "debt": debt_ratio,
            "haram": round(haram_pct, 1),
            "intr": round(interest_ratio, 1),
            "rr": rr,
            "rt": rt,
            "cr": cr,
        }

    except Exception as e:
        print(f"  [ERROR] {ticker}: {e}")
        return None


def run_all():
    """Screen all tickers and write stocks.json."""
    all_tickers = (
        TICKERS_US + TICKERS_DE + TICKERS_SA +
        TICKERS_AE + TICKERS_MY + TICKERS_TR
    )

    print(f"HalalStocks Screener — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"Screening {len(all_tickers)} tickers across 6 exchanges...\n")

    results = []
    failed = []

    for i, ticker in enumerate(all_tickers, 1):
        print(f"[{i:3}/{len(all_tickers)}] {ticker:<12}", end="", flush=True)
        result = screen_ticker(ticker)
        if result:
            results.append(result)
            sts_label = {"h": "ХАЛЯЛЬ", "d": "СОМНИТ.", "r": "ХАРАМ "}.get(result["sts"], "?")
            debt_str = f"{result['debt']:.0f}%" if result["debt"] is not None else "N/A "
            print(f"  → {sts_label}  debt={debt_str:6}  haram={result['haram']:.1f}%")
        else:
            failed.append(ticker)
            print(f"  → FAILED")

        # Rate limiting — be respectful to Yahoo Finance
        time.sleep(0.3)

    # Sort: halal first, then doubtful, then haram; alphabetical within groups
    order = {"h": 0, "d": 1, "r": 2}
    results.sort(key=lambda x: (order.get(x["sts"], 3), x["tk"]))

    # Write output
    output = {
        "meta": {
            "lastUpdated": datetime.now().strftime("%Y-%m-%d"),
            "version": "2.0",
            "dataSource": "yfinance | Автоматический квартальный скрининг",
            "note": "Данные приблизительные. Обновляется автоматически каждый квартал.",
            "totalStocks": len(results),
            "failed": failed,
        },
        "stocks": results,
    }

    out_path = os.path.join(os.path.dirname(__file__), "..", "data", "stocks.json")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✓ Wrote {len(results)} stocks to data/stocks.json")
    print(f"  Halal:     {sum(1 for r in results if r['sts']=='h')}")
    print(f"  Doubtful:  {sum(1 for r in results if r['sts']=='d')}")
    print(f"  Haram:     {sum(1 for r in results if r['sts']=='r')}")
    if failed:
        print(f"  Failed:    {len(failed)} — {', '.join(failed)}")


if __name__ == "__main__":
    run_all()
