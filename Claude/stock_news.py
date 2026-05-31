import yfinance as yf
import requests
from deep_translator import GoogleTranslator
import urllib.parse
import json
import sys
import io
import os
import re
from datetime import datetime
from config import NAVER_CLIENT_ID, NAVER_CLIENT_SECRET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── 설정 ──────────────────────────────────────────────────────────────────────
TICKERS = ["QQQ", "VOO", "O", "TSLA", "GOOG", "KO"]
MAX_YAHOO  = 3
MAX_NAVER  = 2

TICKER_COLORS = {
    "QQQ": "#4f8ef7", "VOO": "#34a853", "O": "#e8710a",
    "TSLA": "#cc0000", "GOOG": "#1a73e8", "KO": "#f4333c",
}
TICKER_NAMES = {
    "QQQ": "Invesco QQQ", "VOO": "Vanguard S&P 500", "O": "Realty Income",
    "TSLA": "Tesla", "GOOG": "Alphabet (Google)", "KO": "Coca-Cola",
}

COMMODITIES = [
    {"symbol": "GC=F", "label": "금 (Gold)",   "unit": "USD/oz",   "color": "#b8860b"},
    {"symbol": "SI=F", "label": "은 (Silver)", "unit": "USD/oz",   "color": "#708090"},
]
FOREX = [
    {"symbol": "USDKRW=X", "label": "달러 (USD)", "unit": "원/달러",  "color": "#2563eb", "multiplier": 1},
    {"symbol": "EURKRW=X", "label": "유로 (EUR)", "unit": "원/유로",  "color": "#7c3aed", "multiplier": 1},
    {"symbol": "JPYKRW=X", "label": "엔화 (JPY)", "unit": "원/100엔","color": "#db2777", "multiplier": 100},
    {"symbol": "CNYKRW=X", "label": "위안 (CNY)", "unit": "원/위안", "color": "#d97706", "multiplier": 1},
]

PERIODS = [
    ("1일",  1),   ("1주",  7),   ("1달",  30),  ("3달",  90),
    ("6달",  180), ("1년",  365), ("2년",  730), ("3년",  1095), ("전체", 0),
]
DEFAULT_DAYS = 30

translator = GoogleTranslator(source="auto", target="ko")

# ── 번역 ──────────────────────────────────────────────────────────────────────
def translate(text):
    if not text:
        return ""
    try:
        return translator.translate(text[:4500])
    except Exception:
        return text

# ── 시세 수집 ─────────────────────────────────────────────────────────────────
def fetch_history(symbol, multiplier=1):
    try:
        hist = yf.Ticker(symbol).history(period="max")
        if hist.empty:
            return [], None, None, None
        closes = [round(float(c) * multiplier, 2) for c in hist["Close"].tolist()]
        dates  = [d.strftime("%Y-%m-%d") for d in hist.index]
        labels = [d.strftime("%m/%d") for d in hist.index]
        last   = closes[-1]
        prev   = closes[-2] if len(closes) >= 2 else last
        change = last - prev
        pct    = (change / prev * 100) if prev else 0
        return list(zip(dates, labels, closes)), last, change, pct
    except Exception:
        return [], None, None, None

def fetch_intraday(symbol, multiplier=1):
    try:
        hist = yf.Ticker(symbol).history(period="1d", interval="1m")
        if hist.empty:
            return []
        closes = [round(float(c) * multiplier, 2) for c in hist["Close"].tolist()]
        times  = hist.index.strftime("%H:%M").tolist()
        return list(zip(times, closes))
    except Exception:
        return []

def fetch_stock_history(symbol):
    return fetch_history(symbol)

# ── 뉴스 수집 ─────────────────────────────────────────────────────────────────
def fetch_yahoo_news(symbol):
    try:
        news = yf.Ticker(symbol).news or []
        result = []
        for item in news[:MAX_YAHOO]:
            content  = item.get("content", {})
            title    = content.get("title",   item.get("title", ""))
            summary  = content.get("summary", "")
            pub_date = content.get("pubDate", "")
            link     = content.get("canonicalUrl", {}).get("url", "") or item.get("link", "")
            if pub_date:
                try:
                    dt = datetime.fromisoformat(pub_date.replace("Z", "+00:00"))
                    pub_date = dt.strftime("%Y-%m-%d %H:%M")
                except Exception:
                    pass
            if len(summary) > 300:
                summary = summary[:300] + "..."
            result.append({"title": title, "summary": summary, "date": pub_date, "link": link, "source": "Yahoo Finance"})
        return result
    except Exception:
        return []

def fetch_korean_news(ticker):
    """네이버 뉴스 검색 API로 한국어 뉴스 수집"""
    try:
        query = urllib.parse.quote(f"{ticker} 주식")
        url   = f"https://openapi.naver.com/v1/search/news.json?query={query}&display={MAX_NAVER}&sort=date"
        headers = {
            "X-Naver-Client-Id":     NAVER_CLIENT_ID,
            "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
        }
        resp = requests.get(url, headers=headers, timeout=10)
        items = resp.json().get("items", [])
        result = []
        for item in items:
            title    = re.sub(r"<[^>]+>", "", item.get("title", ""))
            summary  = re.sub(r"<[^>]+>", "", item.get("description", ""))[:200]
            link     = item.get("originallink") or item.get("link", "")
            pub_date = item.get("pubDate", "")
            if pub_date:
                try:
                    import email.utils
                    dt = datetime(*email.utils.parsedate(pub_date)[:6])
                    pub_date = dt.strftime("%Y-%m-%d %H:%M")
                except Exception:
                    pass
            if title:
                result.append({"title": title, "summary": summary, "date": pub_date, "link": link, "source": "네이버뉴스"})
        return result
    except Exception:
        return []

# ── HTML 헬퍼 ─────────────────────────────────────────────────────────────────
def global_period_buttons_html():
    btns = []
    for label, days in PERIODS:
        active = ' class="active"' if days == DEFAULT_DAYS else ''
        btns.append(f'<button{active} onclick="setAllRanges({days},this)">{label}</button>')
    return '<div class="global-period">' + "".join(btns) + '</div>'

def chart_script(cid, color, history, intraday):
    dates    = json.dumps([d for d, _, _ in history])
    labels   = json.dumps([l for _, l, _ in history])
    prices   = json.dumps([p for _, _, p in history])
    i_labels = json.dumps([t for t, _ in intraday])
    i_prices = json.dumps([p for _, p in intraday])
    return f"""<script>
(function(){{
  var allDates={dates},allLabels={labels},allPrices={prices};
  var intradayLabels={i_labels},intradayPrices={i_prices};
  var ctx=document.getElementById('{cid}').getContext('2d');
  var chart=new Chart(ctx,{{
    type:'line',
    data:{{labels:[],datasets:[{{data:[],borderColor:'{color}',borderWidth:2,pointRadius:0,tension:0.3,fill:true,backgroundColor:'{color}20'}}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{mode:'index',intersect:false}}}},
      scales:{{
        x:{{display:true,ticks:{{maxTicksLimit:6,maxRotation:0,font:{{size:10}}}},grid:{{display:false}}}},
        y:{{display:true,position:'right',ticks:{{maxTicksLimit:4,font:{{size:10}}}},grid:{{color:'#f0f0f0'}}}}
      }},
      animation:false
    }}
  }});
  window.CHARTS=window.CHARTS||{{}};
  window.CHARTS['{cid}']={{chart:chart,allDates:allDates,allLabels:allLabels,allPrices:allPrices,intradayLabels:intradayLabels,intradayPrices:intradayPrices}};
}})();
</script>"""

def change_badge(change, pct):
    if change is None:
        return ""
    sign  = "+" if change >= 0 else ""
    color = "#2ecc71" if change >= 0 else "#e74c3c"
    arrow = "▲" if change >= 0 else "▼"
    return f'<span class="change" style="color:{color}">{arrow} {sign}{change:,.2f} ({sign}{pct:.2f}%)</span>'

# ── 시장 섹션 HTML ─────────────────────────────────────────────────────────────
def market_card_html(cid, label, unit, color, history, intraday, last, change, pct):
    if last is None:
        body = "<p class='no-data'>데이터 없음</p>"
        return f'<div class="mcard"><div class="mcard-head" style="background:{color}"><span class="mlabel">{label}</span><small class="munit">{unit}</small></div><div class="mcard-body">{body}</div></div>'

    price_html = f'<div class="mcard-price"><span class="mprice">{last:,.2f}</span>{change_badge(change, pct)}</div>'
    canvas     = f'<div class="mchart-wrap" style="height:100px"><canvas id="{cid}"></canvas></div>'
    script     = chart_script(cid, color, history, intraday)

    return f"""<div class="mcard">
      <div class="mcard-head" style="background:{color}">
        <span class="mlabel">{label}</span><small class="munit">{unit}</small>
      </div>
      <div class="mcard-body">
        {price_html}
        {canvas}
      </div>
    </div>{script}"""

def build_market_section(commodity_data, forex_data):
    comm_html = ""
    for c, hist, intraday, last, change, pct in commodity_data:
        cid = "m_" + c["symbol"].replace("=","").replace("^","")
        comm_html += market_card_html(cid, c["label"], c["unit"], c["color"], hist, intraday, last, change, pct)

    fx_html = ""
    for f, hist, intraday, last, change, pct in forex_data:
        cid = "m_" + f["symbol"].replace("=","").replace("^","")
        fx_html += market_card_html(cid, f["label"], f["unit"], f["color"], hist, intraday, last, change, pct)

    return f"""<section class="market-section">
      <div class="market-group">
        <h2 class="group-title">귀금속</h2>
        <div class="mcard-grid">{comm_html}</div>
      </div>
      <div class="market-group">
        <h2 class="group-title">환율 (원화 기준)</h2>
        <div class="mcard-grid">{fx_html}</div>
      </div>
    </section>"""

# ── 주식 카드 HTML ────────────────────────────────────────────────────────────
def news_item_html(item, is_korean=False):
    title   = item["title"]
    summary = item["summary"]
    link    = item["link"]
    date    = item["date"]
    source  = item["source"]

    if not is_korean:
        title   = translate(title)
        summary = translate(summary)

    src_badge   = f'<span class="src-badge">{source}</span>'
    title_html  = f'<a href="{link}" target="_blank" class="news-title">{title}</a>' if link else f'<span class="news-title">{title}</span>'
    return f"""<div class="news-item">
      {title_html}
      <div class="news-meta">{src_badge}<span class="news-date">{date}</span></div>
      <div class="news-summary">{summary}</div>
    </div>"""

def build_stock_cards(all_stock):
    cards = ""
    for symbol, hist, intraday, last, change, pct, yahoo_news, naver_news in all_stock:
        color = TICKER_COLORS.get(symbol, "#555")
        name  = TICKER_NAMES.get(symbol, symbol)
        cid   = "s_" + symbol

        price_html = ""
        if last is not None:
            price_html = f"""<div class="price-box">
              <span class="price">${last:,.2f}</span>
              {change_badge(change, pct)}
            </div>"""

        canvas = f'<div class="schart-wrap" style="height:120px"><canvas id="{cid}"></canvas></div>'
        script = chart_script(cid, color, hist, intraday) if hist else ""

        all_news  = naver_news + yahoo_news
        news_html = "".join(news_item_html(n, is_korean=(n["source"] != "Yahoo Finance")) for n in all_news)
        if not news_html:
            news_html = "<p class='no-news'>뉴스 없음</p>"

        cards += f"""<div class="card">
          <div class="card-header" style="background:{color}">
            <div><span class="ticker">{symbol}</span><span class="ticker-name">{name}</span></div>
            <span class="badge">뉴스 {len(all_news)}개</span>
          </div>
          {price_html}
          {canvas}
          {script}
          <div class="card-body">{news_html}</div>
        </div>"""
    return cards

# ── 전체 HTML ────────────────────────────────────────────────────────────────
def build_html(now, market_section, stock_cards):
    date_str    = now.strftime("%Y년 %m월 %d일 %H:%M")
    global_btns = global_period_buttons_html()
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>시장 현황 & 주식 뉴스 - {date_str}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI','Malgun Gothic',sans-serif;background:#f0f2f5;color:#333}}
header{{background:#1a1a2e;color:#fff;padding:16px 28px;display:flex;align-items:center;justify-content:space-between;gap:14px}}
.header-left h1{{font-size:1.25rem;font-weight:700}}
.header-left p{{font-size:0.78rem;color:#aaa;margin-top:2px}}
.global-period{{display:flex;flex-wrap:wrap;gap:4px;justify-content:flex-end}}
.global-period button{{font-size:0.72rem;padding:4px 9px;border:1px solid rgba(255,255,255,.25);border-radius:4px;background:transparent;cursor:pointer;color:rgba(255,255,255,.7);transition:all .15s}}
.global-period button:hover{{border-color:rgba(255,255,255,.6);color:#fff}}
.global-period button.active{{background:#fff;border-color:#fff;color:#1a1a2e;font-weight:700}}

/* 시장 섹션 */
.market-section{{padding:22px 28px 0;display:flex;flex-direction:column;gap:18px}}
.group-title{{font-size:0.78rem;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px}}
.mcard-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px}}
.mcard{{background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.07)}}
.mcard-head{{padding:10px 14px;display:flex;justify-content:space-between;align-items:center}}
.mlabel{{font-size:0.88rem;font-weight:700;color:#fff}}
.munit{{font-size:0.68rem;color:rgba(255,255,255,.8)}}
.mcard-body{{padding:10px 14px 8px}}
.mcard-price{{display:flex;flex-direction:column;gap:2px;margin-bottom:6px}}
.mprice{{font-size:1rem;font-weight:700;color:#1a1a2e}}
.mchart-wrap{{position:relative}}

/* 구분선 */
.section-divider{{margin:22px 28px 0;border:none;border-top:2px solid #e2e8f0}}
.section-label{{padding:18px 28px 6px;font-size:0.78rem;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:1px}}

/* 주식 카드 */
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(400px,1fr));gap:20px;padding:10px 28px 28px}}
.card{{background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.08)}}
.card-header{{padding:13px 18px;display:flex;align-items:center;justify-content:space-between}}
.ticker{{font-size:1.15rem;font-weight:700;color:#fff;letter-spacing:1px;margin-right:8px}}
.ticker-name{{font-size:0.75rem;color:rgba(255,255,255,.75)}}
.badge{{font-size:0.7rem;color:rgba(255,255,255,.85);background:rgba(0,0,0,.15);padding:3px 9px;border-radius:20px}}
.price-box{{display:flex;align-items:center;gap:12px;padding:9px 18px;background:#f8f9fa;border-bottom:1px solid #eee}}
.price{{font-size:1.05rem;font-weight:700;color:#1a1a2e}}
.change{{font-size:0.83rem;font-weight:600}}
.schart-wrap{{position:relative;padding:0 18px}}
.card-body{{padding:10px 18px 14px}}

/* 뉴스 */
.news-item{{padding:12px 0;border-bottom:1px solid #f0f0f0}}
.news-item:last-child{{border-bottom:none}}
a.news-title{{display:block;font-size:0.9rem;font-weight:600;line-height:1.45;color:#1a1a2e;margin-bottom:4px;text-decoration:none}}
a.news-title:hover{{color:#4f8ef7;text-decoration:underline}}
span.news-title{{display:block;font-size:0.9rem;font-weight:600;line-height:1.45;color:#1a1a2e;margin-bottom:4px}}
.news-meta{{display:flex;align-items:center;gap:8px;margin-bottom:5px}}
.src-badge{{font-size:0.68rem;padding:1px 7px;border-radius:3px;background:#f0f0f0;color:#666;font-weight:600}}
.news-date{{font-size:0.72rem;color:#aaa}}
.news-summary{{font-size:0.82rem;color:#555;line-height:1.55}}
.no-news,.no-data{{color:#aaa;font-size:0.85rem;padding:10px 0}}

footer{{text-align:center;padding:18px;font-size:0.75rem;color:#bbb}}
</style>
</head>
<body>
<header>
  <div class="header-left">
    <h1>시장 현황 & 주식 뉴스</h1>
    <p>{date_str} 기준</p>
  </div>
  {global_btns}
</header>

{market_section}

<hr class="section-divider">
<div class="section-label">주식 뉴스</div>
<div class="grid">{stock_cards}</div>

<footer>QQQ · VOO · O · TSLA · GOOG · KO &nbsp;|&nbsp; 금 · 은 · 달러 · 유로 · 엔화 · 위안 &nbsp;|&nbsp; Yahoo Finance + 네이버뉴스</footer>

<script>
function setRange(id, days) {{
  var c = window.CHARTS && window.CHARTS[id];
  if (!c) return;
  var fl = [], fp = [];
  if (days === 1) {{
    fl = c.intradayLabels; fp = c.intradayPrices;
  }} else if (days === 0) {{
    fl = c.allLabels; fp = c.allPrices;
  }} else {{
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    for (var i = 0; i < c.allDates.length; i++) {{
      if (new Date(c.allDates[i]) >= cutoff) {{ fl.push(c.allLabels[i]); fp.push(c.allPrices[i]); }}
    }}
  }}
  c.chart.data.labels = fl;
  c.chart.data.datasets[0].data = fp;
  c.chart.update('none');
}}
function setAllRanges(days, btn) {{
  Object.keys(window.CHARTS || {{}}).forEach(function(id) {{ setRange(id, days); }});
  document.querySelectorAll('.global-period button').forEach(function(b) {{ b.classList.remove('active'); }});
  if (btn) btn.classList.add('active');
}}
window.addEventListener('load', function() {{ setAllRanges({DEFAULT_DAYS}, document.querySelector('.global-period button.active')); }});
</script>
</body>
</html>"""

# ── 텍스트 포맷 ──────────────────────────────────────────────────────────────
def format_text(symbol, yahoo_news, naver_news):
    lines = [f"\n{'='*50}", f"  [{symbol}] 뉴스 (네이버 {len(naver_news)}개 + Yahoo {len(yahoo_news)}개)", f"{'='*50}"]
    for i, n in enumerate(naver_news + yahoo_news, 1):
        lines.append(f"\n  {i}. [{n['source']}] {n['title']}")
        if n["date"]:
            lines.append(f"     날짜: {n['date']}")
        if n["summary"]:
            lines.append(f"     요약: {n['summary']}")
    return "\n".join(lines)

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    now = datetime.now()
    print(f"\n주식 뉴스 요약 - {now.strftime('%Y년 %m월 %d일 %H:%M')}\n", flush=True)

    # 귀금속 + 환율
    print("귀금속 수집 중...", flush=True)
    commodity_data = []
    for c in COMMODITIES:
        hist, last, change, pct = fetch_history(c["symbol"], c.get("multiplier", 1))
        intraday = fetch_intraday(c["symbol"], c.get("multiplier", 1))
        commodity_data.append((c, hist, intraday, last, change, pct))

    print("환율 수집 중...", flush=True)
    forex_data = []
    for f in FOREX:
        hist, last, change, pct = fetch_history(f["symbol"], f.get("multiplier", 1))
        intraday = fetch_intraday(f["symbol"], f.get("multiplier", 1))
        forex_data.append((f, hist, intraday, last, change, pct))

    # 주식
    all_stock = []
    text_parts = []
    for symbol in TICKERS:
        print(f"  {symbol} 수집 중...", flush=True)
        hist, last, change, pct = fetch_stock_history(symbol)
        intraday    = fetch_intraday(symbol)
        yahoo_news  = fetch_yahoo_news(symbol)
        korean_news = fetch_korean_news(symbol)
        all_stock.append((symbol, hist, intraday, last, change, pct, yahoo_news, korean_news))
        text_parts.append(format_text(symbol, yahoo_news, korean_news))

    output = f"\n주식 뉴스 요약 - {now.strftime('%Y년 %m월 %d일 %H:%M')}\n" + "\n".join(text_parts) + f"\n{'='*50}\n"
    print(output)

    result_dir = os.path.join(os.path.dirname(__file__), "news_results")
    os.makedirs(result_dir, exist_ok=True)
    date_str = now.strftime("%Y-%m-%d")

    with open(os.path.join(result_dir, f"{date_str}.txt"), "w", encoding="utf-8") as f:
        f.write(output)

    print("HTML 생성 중...", flush=True)
    market_section = build_market_section(commodity_data, forex_data)
    stock_cards    = build_stock_cards(all_stock)
    html           = build_html(now, market_section, stock_cards)

    html_path = os.path.join(result_dir, f"{date_str}.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"TXT  저장: {os.path.join(result_dir, date_str + '.txt')}")
    print(f"HTML 저장: {html_path}")

if __name__ == "__main__":
    main()
