"""
seo_gap_app.py  —  Streamlit arayüzü
Çalıştırma: streamlit run seo_gap_app.py
"""

import os, re, time
from collections import defaultdict
from datetime import datetime
from io import BytesIO

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="SEO Gap Analyzer", page_icon="🔍", layout="wide")

# ══════════════════════════════════════════════════════════════
# TASARIM — açık tema, veri-yoğun dashboard
# ══════════════════════════════════════════════════════════════

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f7f8fa; color: #1a1d27; }
#MainMenu, header, footer { visibility: hidden; }
.block-container { padding-top: 2rem; max-width: 1280px; }

.app-header {
    border-bottom: 1px solid #e3e5ea; padding-bottom: 16px; margin-bottom: 24px;
}
.app-header h1 { font-size: 1.4rem; font-weight: 700; color: #14161f; margin: 0; letter-spacing: -0.01em; }
.app-header p { color: #767a8a; font-size: 0.85rem; margin: 2px 0 0; }

div[data-testid="stHorizontalBlock"] { align-items: stretch !important; }
.stTextInput input {
    background: #fff !important; border: 1px solid #d8dae2 !important; border-radius: 8px !important;
    color: #14161f !important; font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.92rem !important; height: 44px !important; padding: 0 14px !important;
}
.stTextInput input:focus { border-color: #4f6ef7 !important; box-shadow: 0 0 0 3px rgba(79,110,247,0.12) !important; }
.stButton button {
    background: #4f6ef7 !important; color: #fff !important; border: none !important;
    border-radius: 8px !important; height: 44px !important; font-weight: 600 !important;
    font-size: 0.9rem !important; width: 100% !important; margin-top: 0 !important;
}
.stButton button:hover { background: #3d5be0 !important; }

.metric-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin: 24px 0; }
.metric-card { background: #fff; border: 1px solid #e3e5ea; border-radius: 10px; padding: 16px 18px; }
.metric-card .label { font-size: 0.68rem; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; color: #9296a3; margin-bottom: 6px; }
.metric-card .value { font-size: 1.65rem; font-weight: 700; color: #14161f; font-family: 'JetBrains Mono', monospace; line-height: 1; }
.metric-card .sub { font-size: 0.74rem; color: #9296a3; margin-top: 4px; }
.metric-card.warn .value { color: #d23c3c; }
.metric-card.good .value { color: #1a9b5c; }

.section-title { font-size: 0.78rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: #4a4e5c; margin: 28px 0 12px; display: flex; align-items: center; gap: 8px; }
.section-title .count { background: #eceef3; color: #767a8a; font-size: 0.7rem; padding: 1px 7px; border-radius: 10px; font-weight: 600; }

.legend { display: flex; gap: 18px; margin-bottom: 12px; font-size: 0.78rem; color: #767a8a; }
.legend span.dot { display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 6px; }

.dtable { background: #fff; border: 1px solid #e3e5ea; border-radius: 10px; overflow: hidden; }
.dtable .row { display: grid; padding: 11px 16px; border-bottom: 1px solid #eef0f4; align-items: center; font-size: 0.86rem; }
.dtable .head { background: #fafbfc; font-size: 0.68rem; font-weight: 700; letter-spacing: 0.05em; text-transform: uppercase; color: #9296a3; }
.cat-name { font-weight: 600; color: #14161f; }
.mono { font-family: 'JetBrains Mono', monospace; color: #4a4e5c; font-size: 0.84rem; }
.sample-kw { font-size: 0.74rem; color: #9296a3; padding: 0 16px 11px; border-bottom: 1px solid #eef0f4; }

.badge { display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 0.68rem; font-weight: 700; letter-spacing: 0.02em; }
.badge-opportunity { background: #e3f5ec; color: #1a9b5c; }
.badge-high  { background: #fdf0e0; color: #b8650f; }
.badge-none  { background: #fdeaea; color: #d23c3c; }

div[data-testid="stDownloadButton"] button {
    background: #fff !important; border: 1px solid #d8dae2 !important; color: #4a4e5c !important;
    border-radius: 8px !important; font-weight: 600 !important;
}

.warn-box { background: #fdf0e0; border: 1px solid #f0d4a8; border-radius: 8px; padding: 12px 16px; margin-bottom: 18px; font-size: 0.82rem; color: #8a5a10; }
.stProgress > div > div { background: #4f6ef7 !important; }
.streamlit-expanderHeader { font-size: 0.85rem !important; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# CREDENTIALS
# ══════════════════════════════════════════════════════════════

DFS_LOGIN    = os.getenv("DATAFORSEO_LOGIN", "")
DFS_PASSWORD = os.getenv("DATAFORSEO_PASSWORD", "")

if not DFS_LOGIN or not DFS_PASSWORD:
    st.error("⚠️ `.env` dosyasında DATAFORSEO_LOGIN ve DATAFORSEO_PASSWORD eksik.")
    st.stop()

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SEOGapAnalyzer/1.0)"}


# ══════════════════════════════════════════════════════════════
# 1. SHOPIFY — kategoriler + TÜM ürün başlıkları
# ══════════════════════════════════════════════════════════════

def detect_platform(domain):
    try:
        r = requests.get(f"https://{domain}/collections.json?limit=1", headers=HEADERS, timeout=10)
        if r.status_code == 200 and "collections" in r.json():
            return "shopify"
    except Exception:
        pass
    return "generic"


def fetch_all_collections(domain, progress_cb=None):
    collections, page = [], 1
    while True:
        try:
            r = requests.get(f"https://{domain}/collections.json?limit=250&page={page}", headers=HEADERS, timeout=15)
            cols = r.json().get("collections", [])
        except Exception:
            break
        if not cols:
            break
        collections.extend(cols)
        if progress_cb:
            progress_cb(f"Koleksiyonlar çekiliyor... ({len(collections)} bulundu)")
        if len(cols) < 250:
            break
        page += 1
        time.sleep(0.2)
    return collections


def fetch_all_products(domain, progress_cb=None, max_pages=200):
    products, page = [], 1
    while page <= max_pages:
        try:
            r = requests.get(f"https://{domain}/products.json?limit=250&page={page}", headers=HEADERS, timeout=20)
            items = r.json().get("products", [])
        except Exception:
            break
        if not items:
            break
        for p in items:
            products.append({
                "title": p.get("title", ""),
                "handle": p.get("handle", ""),
            })
        if progress_cb:
            progress_cb(f"Ürünler çekiliyor... ({len(products)} ürün)")
        if len(items) < 250:
            break
        page += 1
        time.sleep(0.15)
    return products


NOISE_PATTERNS = [
    r'^\d+\s*(tl|%)', r'\btest\b', r'\bindirim\w*\b', r'\bkampanya\b',
    r'^\d+[a-z]?$', r'\boutlet\b', r'\bdiscount\b',
    r'^\d+\s*tl\s*(alt[ıi]|üzeri|ve)', r'^#',
]
NOISE_RE = re.compile("|".join(NOISE_PATTERNS), re.IGNORECASE)


def build_categories(collections, progress_cb=None):
    real, noise = [], []
    for c in collections:
        title, handle = c.get("title", ""), c.get("handle", "")
        bucket = noise if (NOISE_RE.search(title) or NOISE_RE.search(handle)) else real
        bucket.append({"title": title, "handle": handle, "product_count": None, "source": "shopify_api"})
    if progress_cb:
        progress_cb(f"{len(real)} gerçek kategori, {len(noise)} kampanya/test koleksiyonu ayrıldı.")
    return real, noise


def fetch_category_product_counts(domain, categories, progress_cb=None):
    for i, cat in enumerate(categories):
        try:
            r = requests.get(f"https://{domain}/collections/{cat['handle']}/products.json?limit=250", headers=HEADERS, timeout=15)
            items = r.json().get("products", [])
            cat["product_count"] = len(items)
        except Exception:
            cat["product_count"] = 0
        if progress_cb and i % 20 == 0:
            progress_cb(f"Kategori ürün sayıları hesaplanıyor... ({i+1}/{len(categories)})")
    return categories


def fetch_site_data(domain, progress_cb=None):
    platform = detect_platform(domain)
    if platform != "shopify":
        if progress_cb:
            progress_cb("⚠️ Shopify olmayan siteler için ürün eşleştirmesi şu an desteklenmiyor.")
        return [], [], []
    if progress_cb:
        progress_cb("Platform: SHOPIFY")
    collections = fetch_all_collections(domain, progress_cb)
    real_cats, noise_cats = build_categories(collections, progress_cb)
    real_cats = fetch_category_product_counts(domain, real_cats, progress_cb)
    products = fetch_all_products(domain, progress_cb)
    return real_cats, noise_cats, products


# ══════════════════════════════════════════════════════════════
# 2. DATAFORSEO
# ══════════════════════════════════════════════════════════════

AUTH = (DFS_LOGIN, DFS_PASSWORD)
BASE = "https://api.dataforseo.com/v3"

def _dfs_post(endpoint, payload, progress_cb=None):
    try:
        r    = requests.post(f"{BASE}/{endpoint}", auth=AUTH, json=payload, timeout=60)
        data = r.json()
        if data.get("status_code") != 20000:
            if progress_cb:
                progress_cb(f"⚠️ DataForSEO hatası: {data.get('status_code')} {data.get('status_message')}")
            return None
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
                "limit": 1000, "filters": ["keyword_data.keyword_info.search_volume", ">", 0],
                "order_by": ["keyword_data.keyword_info.search_volume,desc"]}]
    data = _dfs_post("dataforseo_labs/google/ranked_keywords/live", payload, progress_cb)
    if not data:
        return []
    keywords = []
    try:
        result = data["tasks"][0].get("result")
        if not result or not result[0].get("items"):
            if progress_cb:
                progress_cb("⚠️ DataForSEO sonuç döndü ama keyword listesi boş.")
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
# 3. ANALYZER — keyword × kategori × ürün eşleştirmesi
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

def _matches_category(keyword, cat):
    return bool(_tokens(keyword) & (_tokens(cat["title"]) | _tokens(cat["handle"])))

def count_matching_products(query_tokens, products):
    if not query_tokens:
        return 0
    count = 0
    for p in products:
        p_tokens = _tokens(p["title"])
        if query_tokens & p_tokens and len(query_tokens & p_tokens) >= len(query_tokens):
            count += 1
    return count

def run_gap_analysis(keywords, categories, products, progress_cb=None):
    if progress_cb:
        progress_cb("Gap analizi yapılıyor...")

    matched, orphans = [], []
    groups = defaultdict(lambda: {"keywords":[], "total_volume":0, "modifier":None, "base_cat":None})

    for kw in keywords:
        word, vol = kw.get("keyword",""), kw.get("search_volume",0)
        if vol < 100:
            continue
        if any(_matches_category(word, c) for c in categories):
            matched.append(kw); continue
        mod = next((t for t in _norm(word).split() if t in MODIFIERS), None)
        if mod:
            base = _norm(word).replace(mod,"").strip()[:30]
            key  = f"{mod}__{base}"
            groups[key]["keywords"].append(kw)
            groups[key]["total_volume"] += vol
            groups[key]["modifier"] = mod
            groups[key]["base_cat"] = next((c["title"] for c in categories if _matches_category(base, c)), None)
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
        label = f"{mod.title()} {_norm(top['keyword']).replace(mod,'').strip().title()}".strip()

        query_tokens = _tokens(top["keyword"])
        real_product_count = count_matching_products(query_tokens, products)

        if real_product_count > 0:
            status = "opportunity"
        elif vol >= 500:
            status = "high"
        else:
            status = "none"

        gaps.append({
            "suggested_category":  label,
            "modifier":            mod,
            "base_category":       b_cat or "Belirsiz",
            "keyword_count":       len(group["keywords"]),
            "total_search_volume": vol,
            "top_keyword":         top["keyword"],
            "top_kw_volume":       top["search_volume"],
            "matching_product_count": real_product_count,
            "status":              status,
            "sample_keywords":     ", ".join(k["keyword"] for k in sorted(
                group["keywords"], key=lambda x: x["search_volume"], reverse=True)[:5]),
        })
    gaps.sort(key=lambda x: x["total_search_volume"], reverse=True)
    return {"matched": matched, "gaps": gaps, "orphans": orphans}


# ══════════════════════════════════════════════════════════════
# 4. CHARTS
# ══════════════════════════════════════════════════════════════

STATUS_LABEL = {"opportunity": "ÜRÜN VAR · SAYFA YOK", "high": "HACİM VAR · ÜRÜN YOK", "none": "DÜŞÜK ÖNCELİK"}
STATUS_BADGE_CLASS = {"opportunity": "badge-opportunity", "high": "badge-high", "none": "badge-none"}
STATUS_COLOR = {"opportunity": "#1a9b5c", "high": "#d97a1f", "none": "#c23d3d"}


def build_treemap(categories, gaps, max_categories=20, max_gaps=10):
    top_cats = sorted(categories, key=lambda c: c["product_count"] or 0, reverse=True)[:max_categories]
    top_gaps = gaps[:max_gaps]
    labels, parents, values, colors, texts = ["Site"], [""], [0], ["#eceef3"], [""]

    for cat in top_cats:
        labels.append(cat["title"]); parents.append("Site")
        values.append(max(cat["product_count"] or 0, 5))
        colors.append("#cdeedb")
        texts.append(f"Ürün: {cat['product_count']}")

    for gap in top_gaps:
        labels.append(f"+ {gap['suggested_category']}"); parents.append("Site")
        values.append(max(gap["total_search_volume"] // 5, 30))
        colors.append({"opportunity":"#bfe8d2","high":"#fad9b3","none":"#f7c4c4"}[gap["status"]])
        texts.append(f"Hacim: {gap['total_search_volume']:,} · Eşleşen ürün: {gap['matching_product_count']}")

    fig = go.Figure(go.Treemap(
        labels=labels, parents=parents, values=values, customdata=texts,
        hovertemplate="<b>%{label}</b><br>%{customdata}<extra></extra>",
        marker=dict(colors=colors, line=dict(width=2, color="#fff")),
        textfont=dict(family="Inter", size=13, color="#14161f"),
        pathbar=dict(visible=False),
    ))
    fig.update_layout(margin=dict(l=0,r=0,t=0,b=0), paper_bgcolor="#fff", plot_bgcolor="#fff", height=420)
    return fig


def build_volume_bar(gaps):
    if not gaps:
        return None
    top = gaps[:15]
    fig = go.Figure(go.Bar(
        x=[g["total_search_volume"] for g in top],
        y=[g["suggested_category"] for g in top],
        orientation="h",
        marker_color=[STATUS_COLOR[g["status"]] for g in top],
        text=[f"{g['total_search_volume']:,}" for g in top],
        textposition="outside",
        textfont=dict(family="JetBrains Mono", size=11, color="#4a4e5c"),
        hovertemplate="<b>%{y}</b><br>Hacim: %{x:,}<extra></extra>",
    ))
    fig.update_layout(
        margin=dict(l=0,r=50,t=0,b=0), paper_bgcolor="#fff", plot_bgcolor="#fff",
        height=max(280, len(top)*30),
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
        yaxis=dict(showgrid=False, color="#4a4e5c", tickfont=dict(family="Inter", size=12)),
    )
    return fig


# ══════════════════════════════════════════════════════════════
# 5. EXCEL EXPORT
# ══════════════════════════════════════════════════════════════

def _hdr(ws, r, c, v):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    cell.fill = PatternFill("solid", fgColor="1F3864")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, r, c, v, bg=None, bold=False, align="left"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name="Arial", bold=bold, size=9)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    s = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=s, right=s, top=s, bottom=s)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def generate_excel(gaps, categories, noise_categories, keywords, domain):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    status_label = {"opportunity": "URUN VAR - SAYFA YOK", "high": "HACIM VAR - URUN YOK", "none": "DUSUK ONCELIK"}
    status_color = {"opportunity": "C6EFCE", "high": "FFE2B3", "none": "FFC7CE"}

    ws = wb.create_sheet("Kategori Onerileri")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:J1")
    ws["A1"] = f"SEO Kategori Gap Analizi - {domain}"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:J2")
    ws["A2"] = f"Olusturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=9, color="888888")
    ws["A2"].alignment = Alignment(horizontal="center")
    hdrs = ["Onerilen Kategori","Modifier","Baz Kategori","Keyword Sayisi","Toplam Hacim",
            "En Yuksek KW","En Yuksek KW Hacmi","Eslesen Urun Sayisi","Durum","Ornek Keywordler"]
    widths = [28,15,22,14,14,28,18,16,22,50]
    for c,(h,w) in enumerate(zip(hdrs,widths),1):
        _hdr(ws,3,c,h); ws.column_dimensions[get_column_letter(c)].width = w
    for i,g in enumerate(gaps,4):
        bg = status_color.get(g["status"])
        row = [g["suggested_category"],g["modifier"],g["base_category"],g["keyword_count"],
               g["total_search_volume"],g["top_keyword"],g["top_kw_volume"],g["matching_product_count"],
               status_label.get(g["status"]), g["sample_keywords"]]
        for c,v in enumerate(row,1):
            _cell(ws,i,c,v,bg=bg,bold=(c==9),align="right" if isinstance(v,int) else "left")
    ws.freeze_panes = "A4"

    ws2 = wb.create_sheet("Mevcut Kategoriler")
    ws2.sheet_view.showGridLines = False
    for c,(h,w) in enumerate(zip(["Kategori","Handle","Urun Sayisi"],[28,25,14]),1):
        _hdr(ws2,1,c,h); ws2.column_dimensions[get_column_letter(c)].width = w
    for i,cat in enumerate(categories,2):
        bg = "F2F2F2" if i%2==0 else None
        for c,v in enumerate([cat["title"],cat["handle"],cat["product_count"]],1):
            _cell(ws2,i,c,v,bg=bg,align="right" if isinstance(v,int) else "left")
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("Filtrelenenler")
    ws3.sheet_view.showGridLines = False
    for c,(h,w) in enumerate(zip(["Koleksiyon","Handle"],[28,25]),1):
        _hdr(ws3,1,c,h); ws3.column_dimensions[get_column_letter(c)].width = w
    for i,cat in enumerate(noise_categories,2):
        bg = "F2F2F2" if i%2==0 else None
        for c,v in enumerate([cat["title"],cat["handle"]],1):
            _cell(ws3,i,c,v,bg=bg)
    ws3.freeze_panes = "A2"

    ws4 = wb.create_sheet("Keyword Listesi")
    ws4.sheet_view.showGridLines = False
    for c,(h,w) in enumerate(zip(["Keyword","Hacim","CPC ($)","Rekabet","Pozisyon","URL"],[35,13,10,12,12,55]),1):
        _hdr(ws4,1,c,h); ws4.column_dimensions[get_column_letter(c)].width = w
    for i,kw in enumerate(keywords,2):
        bg = "F2F2F2" if i%2==0 else None
        for c,v in enumerate([kw["keyword"],kw["search_volume"],kw["cpc"],kw["competition"],kw["position"],kw["url"]],1):
            _cell(ws4,i,c,v,bg=bg,align="right" if isinstance(v,(int,float)) else "left")
    ws4.freeze_panes = "A2"

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
    <h1>SEO Gap Analyzer</h1>
    <p>Kategori ağacını analiz et, hacmi olan ama sayfası olmayan fırsatları bul.</p>
</div>
""", unsafe_allow_html=True)

col_input, col_btn = st.columns([5, 1])
with col_input:
    domain_input = st.text_input("", placeholder="derimod.com.tr", label_visibility="collapsed")
with col_btn:
    run_btn = st.button("Analiz Et")

if "results" not in st.session_state:
    st.session_state.results = None

if run_btn and domain_input:
    domain = domain_input.strip().replace("https://","").replace("http://","").strip("/")
    status = st.empty()
    prog   = st.progress(0)
    log_lines = []

    def cb(msg):
        log_lines.append(msg)
        html = "".join(f"<p style='color:{'#d23c3c' if '⚠' in l else '#767a8a'};font-size:0.8rem;margin:2px 0'>{l}</p>" for l in log_lines[-6:])
        status.markdown(html, unsafe_allow_html=True)

    cb("Site verisi çekiliyor (büyük sitelerde birkaç dakika sürebilir)...")
    prog.progress(5)
    real_cats, noise_cats, products = fetch_site_data(domain, cb)

    prog.progress(55)
    keywords = fetch_domain_keywords(domain, cb)

    prog.progress(80)
    results = run_gap_analysis(keywords, real_cats, products, cb)
    results.update({"categories": real_cats, "noise_categories": noise_cats,
                     "products": products, "keywords": keywords, "domain": domain, "logs": log_lines})
    prog.progress(100)

    status.empty(); prog.empty()
    st.session_state.results = results

if st.session_state.results:
    r = st.session_state.results
    gaps, cats, noise_cats = r["gaps"], r["categories"], r.get("noise_categories", [])
    products, kws, dom = r.get("products", []), r["keywords"], r["domain"]

    warnings = [l for l in r.get("logs", []) if "⚠" in l]
    if warnings:
        st.markdown("<div class='warn-box'>" + "<br>".join(warnings) + "</div>", unsafe_allow_html=True)

    opp_count  = sum(1 for g in gaps if g["status"] == "opportunity")
    high_count = sum(1 for g in gaps if g["status"] == "high")

    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card"><div class="label">Mevcut Kategori</div><div class="value">{len(cats)}</div><div class="sub">{len(noise_cats)} kampanya/test ayrıldı</div></div>
        <div class="metric-card"><div class="label">Toplam Ürün</div><div class="value">{len(products):,}</div><div class="sub">sitede taranan</div></div>
        <div class="metric-card"><div class="label">Analiz Edilen Keyword</div><div class="value">{len(kws):,}</div><div class="sub">organik keyword</div></div>
        <div class="metric-card good"><div class="label">Gerçek Fırsat</div><div class="value">{opp_count}</div><div class="sub">ürün var, sayfa yok</div></div>
        <div class="metric-card warn"><div class="label">Ürün Yok</div><div class="value">{high_count}</div><div class="sub">hacim var, ürün yok</div></div>
    </div>
    """, unsafe_allow_html=True)

    col_tree, col_bar = st.columns([3, 2])
    with col_tree:
        st.markdown('<div class="section-title">Kategori Haritası</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="legend">
            <span><span class="dot" style="background:#cdeedb"></span>Mevcut kategori</span>
            <span><span class="dot" style="background:#bfe8d2"></span>Fırsat (ürün var)</span>
            <span><span class="dot" style="background:#fad9b3"></span>Hacim var, ürün yok</span>
        </div>
        """, unsafe_allow_html=True)
        if cats or gaps:
            st.plotly_chart(build_treemap(cats, gaps), use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Gösterilecek kategori verisi yok.")

    with col_bar:
        st.markdown('<div class="section-title">Gap Hacimleri</div>', unsafe_allow_html=True)
        bar = build_volume_bar(gaps)
        if bar:
            st.plotly_chart(bar, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Gap önerisi bulunamadı.")

    st.markdown(f'<div class="section-title">Kategori Önerileri <span class="count">{len(gaps)}</span></div>', unsafe_allow_html=True)
    if gaps:
        rows_html = "<div class='dtable'><div class='row head' style='grid-template-columns:2fr 1.3fr 1fr 1fr 1.4fr'>" \
                    "<span>Önerilen Kategori</span><span>Baz Kategori</span><span>Toplam Hacim</span>" \
                    "<span>Eşleşen Ürün</span><span>Durum</span></div>"
        for g in gaps[:60]:
            badge = f'<span class="badge {STATUS_BADGE_CLASS[g["status"]]}">{STATUS_LABEL[g["status"]]}</span>'
            rows_html += (
                f"<div class='row' style='grid-template-columns:2fr 1.3fr 1fr 1fr 1.4fr'>"
                f"<span class='cat-name'>+ {g['suggested_category']}</span>"
                f"<span class='mono'>{g['base_category']}</span>"
                f"<span class='mono'>{g['total_search_volume']:,}</span>"
                f"<span class='mono'>{g['matching_product_count']}</span>"
                f"<span>{badge}</span></div>"
                f"<div class='sample-kw'>{g['sample_keywords']}</div>"
            )
        rows_html += "</div>"
        st.markdown(rows_html, unsafe_allow_html=True)
        if len(gaps) > 60:
            st.caption(f"İlk 60 öneri gösteriliyor. Tam liste ({len(gaps)} satır) Excel raporunda.")
    else:
        st.info("Hiç gap önerisi bulunamadı.")

    if noise_cats:
        with st.expander(f"Filtrelenen kampanya/test koleksiyonları ({len(noise_cats)})"):
            st.write(", ".join(c["title"] for c in noise_cats[:100]))

    st.markdown("<div style='margin-top:20px'>", unsafe_allow_html=True)
    excel_buf = generate_excel(gaps, cats, noise_cats, kws, dom)
    st.download_button(
        label="⬇ Excel Raporu İndir",
        data=excel_buf,
        file_name=f"seo_gap_{dom.replace('.','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("</div>", unsafe_allow_html=True)
