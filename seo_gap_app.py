"""
seo_gap_app.py  —  Streamlit arayüzü
Çalıştırma: streamlit run seo_gap_app.py
"""

import os, re, sys, time
from collections import defaultdict
from datetime import datetime
from urllib.parse import urlparse
import xml.etree.ElementTree as ET
from io import BytesIO

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ── Sayfa ayarları ───────────────────────────────────────────
st.set_page_config(
    page_title="SEO Gap Analyzer",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background: #0f1117; color: #e8eaf0; }

.hero {
    padding: 48px 0 32px;
    text-align: center;
}
.hero h1 {
    font-size: 2.6rem;
    font-weight: 600;
    letter-spacing: -0.03em;
    color: #f0f2f8;
    margin: 0 0 8px;
}
.hero p {
    color: #7c8199;
    font-size: 1rem;
    font-weight: 300;
    margin: 0;
}

.input-row {
    display: flex;
    gap: 12px;
    max-width: 600px;
    margin: 32px auto 0;
}

/* metric cards */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin: 32px 0 24px;
}
.metric-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 12px;
    padding: 20px 24px;
}
.metric-card .label {
    font-size: 0.75rem;
    font-weight: 500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #6b7080;
    margin-bottom: 8px;
}
.metric-card .value {
    font-size: 2rem;
    font-weight: 600;
    color: #f0f2f8;
    font-family: 'DM Mono', monospace;
    line-height: 1;
}
.metric-card .sub {
    font-size: 0.78rem;
    color: #6b7080;
    margin-top: 4px;
}

/* gap table */
.gap-row {
    display: grid;
    grid-template-columns: 2fr 1.2fr 1fr 1fr 1.2fr;
    gap: 0;
    padding: 14px 20px;
    border-bottom: 1px solid #1e2130;
    align-items: center;
    transition: background 0.15s;
}
.gap-row:hover { background: #1a1d27; }
.gap-header {
    background: #13151f;
    border-radius: 10px 10px 0 0;
    border: 1px solid #2a2d3a;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    color: #5a5f75;
}
.gap-body {
    border: 1px solid #2a2d3a;
    border-top: none;
    border-radius: 0 0 10px 10px;
    overflow: hidden;
}
.cat-name { font-weight: 500; color: #dde0ec; font-size: 0.92rem; }
.vol { font-family: 'DM Mono', monospace; font-size: 0.88rem; color: #a0a5bc; }
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.04em;
}
.badge-high  { background: #1a3a2a; color: #4ade80; border: 1px solid #2a5a3a; }
.badge-mid   { background: #2a2a10; color: #facc15; border: 1px solid #4a4a20; }
.badge-warn  { background: #3a1a1a; color: #f87171; border: 1px solid #5a2a2a; }
.badge-none  { background: #1a1d27; color: #6b7080; border: 1px solid #2a2d3a; }

.section-title {
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #5a5f75;
    margin: 32px 0 14px;
}

.stButton > button {
    background: #4f6ef7 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 28px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.95rem !important;
    transition: all 0.2s !important;
    width: 100% !important;
}
.stButton > button:hover { background: #3a57e0 !important; transform: translateY(-1px); }

.stTextInput > div > div > input {
    background: #1a1d27 !important;
    border: 1px solid #2a2d3a !important;
    border-radius: 8px !important;
    color: #e8eaf0 !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.95rem !important;
    padding: 10px 16px !important;
}
.stTextInput > div > div > input:focus {
    border-color: #4f6ef7 !important;
    box-shadow: 0 0 0 2px rgba(79,110,247,0.2) !important;
}

div[data-testid="stDownloadButton"] > button {
    background: #1a1d27 !important;
    border: 1px solid #2a2d3a !important;
    color: #a0a5bc !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
}

.stProgress > div > div { background: #4f6ef7 !important; }
.stSpinner { color: #4f6ef7 !important; }

/* plotly bg fix */
.js-plotly-plot { border-radius: 12px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# CREDENTIALS
# ══════════════════════════════════════════════════════════════

DFS_LOGIN    = os.getenv("DATAFORSEO_LOGIN", "")
DFS_PASSWORD = os.getenv("DATAFORSEO_PASSWORD", "")

if not DFS_LOGIN or not DFS_PASSWORD:
    st.error("⚠️  `.env` dosyasında DATAFORSEO_LOGIN ve DATAFORSEO_PASSWORD eksik.")
    st.stop()


# ══════════════════════════════════════════════════════════════
# CRAWLER
# ══════════════════════════════════════════════════════════════

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SEOGapAnalyzer/1.0)"}

def detect_platform(domain):
    try:
        r = requests.get(f"https://{domain}/collections.json?limit=1", headers=HEADERS, timeout=10)
        if r.status_code == 200 and "collections" in r.json():
            return "shopify"
    except Exception:
        pass
    return "generic"

def get_shopify_categories(domain, progress_cb=None):
    categories, page = [], 1
    while True:
        try:
            r = requests.get(f"https://{domain}/collections.json?limit=250&page={page}", headers=HEADERS, timeout=15)
            cols = r.json().get("collections", [])
        except Exception as e:
            break
        if not cols:
            break
        for col in cols:
            handle = col.get("handle", "")
            count  = _shopify_count(domain, handle)
            categories.append({"title": col.get("title",""), "handle": handle,
                                "url": f"https://{domain}/collections/{handle}",
                                "product_count": count, "source": "shopify_api"})
            if progress_cb:
                progress_cb(f"Kategori: {col.get('title','')} ({count} ürün)")
        if len(cols) < 250:
            break
        page += 1; time.sleep(0.3)
    return categories

def _shopify_count(domain, handle):
    try:
        r = requests.get(f"https://{domain}/collections/{handle}", headers=HEADERS, timeout=10)
        for pat in [r'"products_count"\s*:\s*(\d+)', r'(\d+)\s*(ürün|products?)']:
            m = re.search(pat, r.text, re.IGNORECASE)
            if m:
                return int(m.group(1))
        soup = BeautifulSoup(r.text, "html.parser")
        return len(soup.select(".product-item,.grid__item,[data-product-id]"))
    except Exception:
        return 0

def get_generic_categories(domain, progress_cb=None):
    sitemaps = _discover_sitemaps(domain)
    cat_urls = []
    for sm in sitemaps:
        for u in _parse_sitemap(sm):
            path = urlparse(u).path.lower()
            if any(s in path for s in ["/kategori","/category","/collections","/urun-listesi","/c/","/cat/","/shop/"]):
                cat_urls.append(u)
    cat_urls = list(set(cat_urls))
    categories = []
    for url in cat_urls[:100]:
        info = _scrape_category(url)
        if info:
            categories.append(info)
            if progress_cb:
                progress_cb(f"Kategori: {info['title']} ({info['product_count']} ürün)")
        time.sleep(0.2)
    return categories

def _discover_sitemaps(domain):
    candidates = [f"https://{domain}/sitemap.xml", f"https://{domain}/sitemap_index.xml"]
    try:
        r = requests.get(f"https://{domain}/robots.txt", headers=HEADERS, timeout=10)
        for line in r.text.splitlines():
            if line.lower().startswith("sitemap:"):
                candidates.append(line.split(":",1)[1].strip())
    except Exception:
        pass
    return [u for u in dict.fromkeys(candidates) if _url_ok(u)]

def _url_ok(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        return r.status_code == 200 and ("<url>" in r.text or "<sitemap>" in r.text)
    except Exception:
        return False

def _parse_sitemap(url):
    urls = []
    try:
        r    = requests.get(url, headers=HEADERS, timeout=15)
        root = ET.fromstring(r.content)
        ns   = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        for loc in root.findall("sm:sitemap/sm:loc", ns):
            urls.extend(_parse_sitemap(loc.text.strip()))
        for loc in root.findall("sm:url/sm:loc", ns):
            urls.append(loc.text.strip())
    except Exception:
        pass
    return urls

def _scrape_category(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        soup  = BeautifulSoup(r.text, "html.parser")
        title = soup.find("h1").get_text(strip=True) if soup.find("h1") else urlparse(url).path.split("/")[-1].replace("-"," ").title()
        count = 0
        for pat in [r'(\d+)\s*(ürün|products?|sonuç|results?)']:
            m = re.search(pat, r.text, re.IGNORECASE)
            if m:
                count = int(m.group(1)); break
        if not count:
            count = len(soup.select(".product-item,.product-card,[data-product-id],article.product"))
        return {"title": title, "handle": urlparse(url).path.strip("/").split("/")[-1],
                "url": url, "product_count": count, "source": "sitemap_scrape"}
    except Exception:
        return None

def fetch_site_categories(domain, progress_cb=None):
    platform = detect_platform(domain)
    if progress_cb:
        progress_cb(f"Platform tespit edildi: {platform.upper()}")
    return get_shopify_categories(domain, progress_cb) if platform == "shopify" else get_generic_categories(domain, progress_cb)


# ══════════════════════════════════════════════════════════════
# DATAFORSEO
# ══════════════════════════════════════════════════════════════

AUTH = (DFS_LOGIN, DFS_PASSWORD)
BASE = "https://api.dataforseo.com/v3"

def _dfs_post(endpoint, payload, progress_cb=None):
    try:
        r    = requests.post(f"{BASE}/{endpoint}", auth=AUTH, json=payload, timeout=60)
        data = r.json()
        if data.get("status_code") != 20000:
            if progress_cb:
                progress_cb(f"⚠️ DataForSEO hatası (üst seviye): {data.get('status_code')} {data.get('status_message')}")
            return None
        # Task seviyesinde de hata olabilir, üst seviye 20000 olsa bile
        task = data.get("tasks", [{}])[0]
        if task.get("status_code") != 20000:
            if progress_cb:
                progress_cb(f"⚠️ DataForSEO task hatası: {task.get('status_code')} {task.get('status_message')}")
            return None
        return data
    except Exception as e:
        if progress_cb:
            progress_cb(f"⚠️ İstek hatası: {e}")
        return None

def fetch_domain_keywords(domain, progress_cb=None):
    if progress_cb:
        progress_cb("DataForSEO'dan keyword'ler çekiliyor...")
    payload = [{"target": domain, "language_code": "tr", "location_code": 2792,
                "limit": 2000, "filters": ["keyword_data.keyword_info.search_volume", ">", 0],
                "order_by": ["keyword_data.keyword_info.search_volume,desc"]}]
    data = _dfs_post("dataforseo_labs/google/ranked_keywords/live", payload, progress_cb)
    if not data:
        return []
    keywords = []
    try:
        result = data["tasks"][0].get("result")
        if not result or not result[0].get("items"):
            if progress_cb:
                progress_cb("⚠️ DataForSEO sonuç döndü ama içinde keyword yok (items boş).")
            return []
        for item in result[0]["items"]:
            kw   = item.get("keyword_data", {})
            info = kw.get("keyword_info", {})
            serp = item.get("ranked_serp_element", {}).get("serp_element", {})
            keywords.append({"keyword": kw.get("keyword",""), "search_volume": info.get("search_volume",0),
                             "cpc": info.get("cpc",0), "competition": info.get("competition",0),
                             "position": serp.get("rank_group",0), "url": serp.get("url","")})
        if progress_cb:
            progress_cb(f"{len(keywords)} keyword çekildi.")
    except Exception as e:
        if progress_cb:
            progress_cb(f"⚠️ Parse hatası: {e}")
    return keywords


# ══════════════════════════════════════════════════════════════
# ANALYZER
# ══════════════════════════════════════════════════════════════

STOP = {"ve","ile","icin","bir","bu","da","de","mi","mu","the","and","for","in","of","a","an","on","at","to"}
TR_NORM = str.maketrans("ığüşöçİĞÜŞÖÇ","igusocIGUSOC")
MODIFIERS = set([
    "siyah","beyaz","kirmizi","mavi","yesil","sari","mor","pembe","turuncu","gri",
    "kahverengi","bej","lacivert","bordo","krem","ekru","haki","black","white","red",
    "blue","green","yellow","grey","brown","beige","navy","gumus","altin",
    "deri","suet","tokali","bagcikli","fermuarli","platform","topuklu","duz","kalin",
    "ince","uzun","kisa","maxi","mini","midi","spor","klasik","casual","formal",
    "gunluk","yazlik","kislik","sik","rahat",
])

def _norm(text):
    return re.sub(r"[^a-z0-9\s]"," ", text.lower().translate(TR_NORM)).strip()

def _tokens(text):
    return set(_norm(text).split()) - STOP

def _matches(keyword, cat):
    return bool(_tokens(keyword) & (_tokens(cat["title"]) | _tokens(cat["handle"])))

def run_gap_analysis(keywords, categories, progress_cb=None):
    if progress_cb:
        progress_cb("Gap analizi yapılıyor...")
    cat_product_map = {_norm(c["title"]): c.get("product_count",0) for c in categories}
    matched, orphans = [], []
    groups = defaultdict(lambda: {"keywords":[], "total_volume":0, "modifier":None, "base_cat":None})

    for kw in keywords:
        word, vol = kw.get("keyword",""), kw.get("search_volume",0)
        if vol < 100:
            continue
        if any(_matches(word, c) for c in categories):
            matched.append(kw); continue
        mod = next((t for t in _norm(word).split() if t in MODIFIERS), None)
        if mod:
            base = _norm(word).replace(mod,"").strip()[:30]
            key  = f"{mod}__{base}"
            groups[key]["keywords"].append(kw)
            groups[key]["total_volume"] += vol
            groups[key]["modifier"] = mod
            groups[key]["base_cat"] = next((c["title"] for c in categories if _matches(base, c)), None)
        else:
            orphans.append(kw)

    gaps = []
    for group in groups.values():
        if not group["keywords"]:
            continue
        top   = max(group["keywords"], key=lambda x: x["search_volume"])
        mod   = group["modifier"]
        b_cat = group["base_cat"]
        vol   = group["total_volume"]
        p_cnt = cat_product_map.get(_norm(b_cat), 0) if b_cat else 0
        label = f"{mod.title()} {_norm(top['keyword']).replace(mod,'').strip().title()}".strip()
        prio  = "high" if vol >= 500 else "mid"
        if p_cnt == 0:
            prio = "warn"
        gaps.append({
            "suggested_category":  label,
            "modifier":            mod,
            "base_category":       b_cat or "Belirsiz",
            "keyword_count":       len(group["keywords"]),
            "total_search_volume": vol,
            "top_keyword":         top["keyword"],
            "top_kw_volume":       top["search_volume"],
            "product_count":       p_cnt,
            "priority":            prio,
            "sample_keywords":     ", ".join(k["keyword"] for k in sorted(
                group["keywords"], key=lambda x: x["search_volume"], reverse=True)[:5]),
        })
    gaps.sort(key=lambda x: x["total_search_volume"], reverse=True)
    return {"matched": matched, "gaps": gaps, "orphans": orphans}


# ══════════════════════════════════════════════════════════════
# CHARTS
# ══════════════════════════════════════════════════════════════

def build_treemap(categories, gaps):
    """Mevcut kategoriler + gap önerileri — interaktif treemap."""
    labels, parents, values, colors, texts = ["Site"], [""], [0], ["#1a1d27"], [""]

    for cat in categories:
        labels.append(cat["title"])
        parents.append("Site")
        values.append(max(cat["product_count"], 10))
        colors.append("#1e3a2a")   # koyu yeşil — mevcut
        texts.append(f"Ürün: {cat['product_count']}")

    for gap in gaps:
        labels.append(f"+ {gap['suggested_category']}")
        parents.append("Site")
        values.append(max(gap["total_search_volume"] // 10, 10))
        colors.append("#3a1a1a" if gap["priority"] == "warn" else "#2a2a10")
        texts.append(f"Hacim: {gap['total_search_volume']:,} | Ürün: {gap['product_count']}")

    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=parents,
        values=values,
        customdata=texts,
        hovertemplate="<b>%{label}</b><br>%{customdata}<extra></extra>",
        marker=dict(
            colors=colors,
            line=dict(width=2, color="#0f1117"),
        ),
        textfont=dict(family="DM Sans", size=13, color="#e8eaf0"),
        pathbar=dict(visible=False),
    ))
    fig.update_layout(
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="#0f1117",
        plot_bgcolor="#0f1117",
        height=480,
    )
    return fig


def build_volume_bar(gaps):
    """Gap önerilerinin arama hacmi bar chart'ı."""
    if not gaps:
        return None
    top = gaps[:15]
    color_map = {"high": "#4ade80", "mid": "#facc15", "warn": "#f87171"}
    fig = go.Figure(go.Bar(
        x=[g["total_search_volume"] for g in top],
        y=[g["suggested_category"] for g in top],
        orientation="h",
        marker_color=[color_map.get(g["priority"], "#6b7080") for g in top],
        text=[f"{g['total_search_volume']:,}" for g in top],
        textposition="outside",
        textfont=dict(family="DM Mono", size=11, color="#a0a5bc"),
        hovertemplate="<b>%{y}</b><br>Hacim: %{x:,}<extra></extra>",
    ))
    fig.update_layout(
        margin=dict(l=0, r=60, t=0, b=0),
        paper_bgcolor="#0f1117",
        plot_bgcolor="#0f1117",
        height=max(300, len(top) * 30),
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False,
                   color="#5a5f75"),
        yaxis=dict(showgrid=False, color="#a0a5bc",
                   tickfont=dict(family="DM Sans", size=12)),
    )
    return fig


# ══════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════

def _hdr(ws, r, c, v):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    cell.fill      = PatternFill("solid", fgColor="1F3864")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s = Side(style="thin", color="CCCCCC")
    cell.border    = Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, r, c, v, bg=None, bold=False, align="left"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=bold, size=9)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    s = Side(style="thin", color="CCCCCC")
    cell.border    = Border(left=s, right=s, top=s, bottom=s)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def generate_excel(gaps, categories, keywords, domain):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Sheet 1
    ws = wb.create_sheet("Kategori Önerileri")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:J1")
    ws["A1"].value     = f"SEO Kategori Gap Analizi — {domain}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:J2")
    ws["A2"].value     = f"Oluşturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A2"].font      = Font(name="Arial", size=9, color="888888")
    ws["A2"].alignment = Alignment(horizontal="center")
    hdrs   = ["Önerilen Kategori","Modifier","Baz Kategori","Keyword Sayısı","Toplam Hacim",
              "En Yüksek KW","En Yüksek KW Hacmi","Ürün Sayısı","Öncelik","Örnek Keyword'ler"]
    widths = [28,15,22,14,14,28,18,12,22,50]
    prio_label = {"high": "YÜKSEK", "mid": "ORTA", "warn": "YÜKSEK ⚠️ ÜRÜN YOK"}
    prio_color = {"high": "C6EFCE", "mid": "FFEB9C", "warn": "FFC7CE"}
    for c, (h, w) in enumerate(zip(hdrs, widths), 1):
        _hdr(ws, 3, c, h)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[3].height = 30
    for i, g in enumerate(gaps, 4):
        bg = prio_color.get(g["priority"])
        row_data = [g["suggested_category"], g["modifier"], g["base_category"],
                    g["keyword_count"], g["total_search_volume"], g["top_keyword"],
                    g["top_kw_volume"], g["product_count"],
                    prio_label.get(g["priority"], g["priority"]), g["sample_keywords"]]
        for c, v in enumerate(row_data, 1):
            _cell(ws, i, c, v, bg=bg, bold=(c==9), align="right" if isinstance(v,int) else "left")
        ws.row_dimensions[i].height = 17
    ws.freeze_panes = "A4"

    # Sheet 2
    ws2 = wb.create_sheet("Mevcut Kategoriler")
    ws2.sheet_view.showGridLines = False
    for c, (h, w) in enumerate(zip(["Kategori","Handle","URL","Ürün Sayısı","Kaynak"],
                                    [28,25,50,13,18]), 1):
        _hdr(ws2, 1, c, h)
        ws2.column_dimensions[get_column_letter(c)].width = w
    for i, cat in enumerate(categories, 2):
        bg = "F2F2F2" if i%2==0 else None
        for c, v in enumerate([cat["title"],cat["handle"],cat.get("url",""),cat["product_count"],cat.get("source","")], 1):
            _cell(ws2, i, c, v, bg=bg, align="right" if isinstance(v,int) else "left")
        ws2.row_dimensions[i].height = 15
    ws2.freeze_panes = "A2"

    # Sheet 3
    ws3 = wb.create_sheet("Keyword Listesi")
    ws3.sheet_view.showGridLines = False
    for c, (h, w) in enumerate(zip(["Keyword","Hacim","CPC ($)","Rekabet","Pozisyon","URL"],
                                    [35,13,10,12,12,55]), 1):
        _hdr(ws3, 1, c, h)
        ws3.column_dimensions[get_column_letter(c)].width = w
    for i, kw in enumerate(keywords, 2):
        bg = "F2F2F2" if i%2==0 else None
        for c, v in enumerate([kw["keyword"],kw["search_volume"],kw["cpc"],
                                kw["competition"],kw["position"],kw["url"]], 1):
            _cell(ws3, i, c, v, bg=bg, align="right" if isinstance(v,(int,float)) else "left")
        ws3.row_dimensions[i].height = 15
    ws3.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

st.markdown("""
<div class="hero">
    <h1>SEO Gap Analyzer</h1>
    <p>Kategori ağacını analiz et, eksik alt kategorileri keşfet.</p>
</div>
""", unsafe_allow_html=True)

col_input, col_btn = st.columns([4, 1])
with col_input:
    domain_input = st.text_input("", placeholder="derimod.com.tr", label_visibility="collapsed")
with col_btn:
    st.markdown("<div style='margin-top:4px'>", unsafe_allow_html=True)
    run_btn = st.button("Analiz Et")
    st.markdown("</div>", unsafe_allow_html=True)

# ── State ────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = None

if run_btn and domain_input:
    domain = domain_input.strip().replace("https://","").replace("http://","").strip("/")
    status = st.empty()
    prog   = st.progress(0)
    log_lines = []

    def cb(msg):
        log_lines.append(msg)
        html = "".join(f"<p style='color:{'#f87171' if '⚠️' in l else '#6b7080'};font-size:0.82rem;margin:2px 0'>{l}</p>" for l in log_lines[-8:])
        status.markdown(html, unsafe_allow_html=True)

    with st.spinner(""):
        cb("Kategori ağacı çekiliyor...")
        prog.progress(10)
        categories = fetch_site_categories(domain, cb)

        prog.progress(40)
        cb("Keyword verileri çekiliyor...")
        keywords = fetch_domain_keywords(domain, cb)

        prog.progress(75)
        cb("Gap analizi yapılıyor...")
        results = run_gap_analysis(keywords, categories, cb)
        results["categories"] = categories
        results["keywords"]   = keywords
        results["domain"]     = domain

        prog.progress(100)

    status.empty()
    prog.empty()
    results["logs"] = log_lines
    st.session_state.results = results

# ── Sonuçlar ─────────────────────────────────────────────────
if st.session_state.results:
    r    = st.session_state.results
    gaps = r["gaps"]
    cats = r["categories"]
    kws  = r["keywords"]
    dom  = r["domain"]

    warnings = [l for l in r.get("logs", []) if "⚠️" in l]
    if warnings:
        st.markdown(
            "<div style='background:#3a1a1a;border:1px solid #5a2a2a;border-radius:8px;padding:14px 18px;margin-bottom:20px'>"
            + "".join(f"<p style='color:#f87171;font-size:0.85rem;margin:2px 0'>{w}</p>" for w in warnings)
            + "</div>",
            unsafe_allow_html=True
        )

    # Metrikler
    high_count = sum(1 for g in gaps if g["priority"] == "high")
    warn_count = sum(1 for g in gaps if g["priority"] == "warn")

    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card">
            <div class="label">Mevcut Kategori</div>
            <div class="value">{len(cats)}</div>
            <div class="sub">siteden çekildi</div>
        </div>
        <div class="metric-card">
            <div class="label">Analiz Edilen Keyword</div>
            <div class="value">{len(kws):,}</div>
            <div class="sub">organik keyword</div>
        </div>
        <div class="metric-card">
            <div class="label">Gap Önerisi</div>
            <div class="value">{len(gaps)}</div>
            <div class="sub">{high_count} yüksek öncelikli</div>
        </div>
        <div class="metric-card">
            <div class="label">Ürün Yok Uyarısı</div>
            <div class="value" style="color:#f87171">{warn_count}</div>
            <div class="sub">hacim var, ürün yok</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Treemap + Bar
    col_tree, col_bar = st.columns([3, 2])
    with col_tree:
        st.markdown('<div class="section-title">Kategori Haritası</div>', unsafe_allow_html=True)
        st.markdown("""
        <div style='display:flex;gap:16px;margin-bottom:12px;font-size:0.75rem;color:#6b7080'>
            <span><span style='display:inline-block;width:10px;height:10px;background:#1e3a2a;border-radius:2px;margin-right:6px'></span>Mevcut kategori</span>
            <span><span style='display:inline-block;width:10px;height:10px;background:#2a2a10;border-radius:2px;margin-right:6px'></span>Önerilen</span>
            <span><span style='display:inline-block;width:10px;height:10px;background:#3a1a1a;border-radius:2px;margin-right:6px'></span>Ürün yok</span>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(build_treemap(cats, gaps), use_container_width=True, config={"displayModeBar": False})

    with col_bar:
        st.markdown('<div class="section-title">Gap Hacimleri</div>', unsafe_allow_html=True)
        bar = build_volume_bar(gaps)
        if bar:
            st.plotly_chart(bar, use_container_width=True, config={"displayModeBar": False})

    # Gap tablosu
    st.markdown('<div class="section-title">Kategori Önerileri</div>', unsafe_allow_html=True)
    badge_map = {
        "high": '<span class="badge badge-high">YÜKSEK</span>',
        "mid":  '<span class="badge badge-mid">ORTA</span>',
        "warn": '<span class="badge badge-warn">ÜRÜN YOK</span>',
    }
    st.markdown("""
    <div class="gap-row gap-header">
        <span>Önerilen Kategori</span>
        <span>Baz Kategori</span>
        <span>Toplam Hacim</span>
        <span>Ürün Sayısı</span>
        <span>Öncelik</span>
    </div>
    <div class="gap-body">
    """ + "".join(f"""
        <div class="gap-row">
            <span class="cat-name">+ {g['suggested_category']}</span>
            <span class="vol">{g['base_category']}</span>
            <span class="vol">{g['total_search_volume']:,}</span>
            <span class="vol">{g['product_count']}</span>
            <span>{badge_map.get(g['priority'],'')}</span>
        </div>
    """ for g in gaps) + "</div>", unsafe_allow_html=True)

    # Export
    st.markdown("<div style='margin-top:24px'>", unsafe_allow_html=True)
    excel_buf = generate_excel(gaps, cats, kws, dom)
    st.download_button(
        label="⬇ Excel Raporu İndir",
        data=excel_buf,
        file_name=f"seo_gap_{dom.replace('.','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("</div>", unsafe_allow_html=True)
