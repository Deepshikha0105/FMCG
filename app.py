"""
FMCG M&A Intelligence Newsletter Agent
=======================================
Pipeline: Ingest -> Deduplicate -> Score Relevance -> Score Credibility -> Generate Newsletter

Sources  : NewsAPI (optional) + RSS feeds
Dedup    : URL hash + fuzzy title match (rapidfuzz) + TF-IDF similarity (sklearn, no model download)
Scoring  : Keyword score + LLM binary filter + tiered source credibility
Output   : Streamlit UI + Word / Excel / JSON / CSV downloads
LLM      : Groq API (llama-3.3-70b-versatile)
"""

import streamlit as st
import requests
import feedparser
import hashlib
import json
import re
import io
import time
from datetime import datetime, timedelta
from dataclasses import dataclass, asdict
import pandas as pd
from rapidfuzz import fuzz
from groq import Groq
from urllib.parse import urlparse
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

# ── Optional exports (graceful fallback if not installed) ─────────────────────

try:
    from docx import Document
    from docx.shared import Pt
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    XLSX_AVAILABLE = True
except Exception:
    XLSX_AVAILABLE = False

# =============================================================================
# DATA MODEL
# =============================================================================

@dataclass
class Article:
    id: str = ""
    title: str = ""
    description: str = ""
    url: str = ""
    source_name: str = ""
    source_domain: str = ""
    published_at: str = ""
    query_used: str = ""
    # Scoring fields - populated during pipeline
    keyword_score: float = 0.0
    llm_relevant: bool = False
    deal_type: str = "unknown"
    credibility_score: float = 0.0
    credibility_tier: str = "unverified"
    composite_score: float = 0.0
    llm_reason: str = ""

# =============================================================================
# CONSTANTS
# =============================================================================


FMCG_QUERIES = [
    # Core FMCG M&A
    "FMCG acquisition merger deal",
    "consumer goods company acquisition",
    "consumer goods merger deal",
    "FMCG company buyout takeover",
    "CPG acquisition merger investment",
    "packaged goods company acquisition deal",
    "consumer staples merger joint venture",
    "FMCG private equity investment deal",
    # Category specific
    "food company acquisition merger",
    "food beverage brand acquisition investment",
    "beverage company merger deal",
    "snack brand acquisition buyout",
    "dairy company acquisition merger",
    "confectionery brand acquisition deal",
    "alcohol spirits brand acquisition",
    "beer wine spirits merger acquisition",
    "personal care brand buyout stake",
    "beauty brand acquisition merger",
    "skincare brand acquisition investment",
    "haircare brand buyout deal",
    "cosmetics company merger acquisition",
    "household products company acquisition",
    "home care brand merger deal",
    "oral care brand acquisition",
    "baby care brand acquisition deal",
    "pet food company acquisition merger",
    "nutrition health brand acquisition",
    "supplement brand acquisition deal",
    "organic food brand acquisition",
    "plant based food company merger",
    # India specific
    "HUL Nestle Dabur Marico ITC acquisition",
    "Tata Consumer Godrej Emami acquisition",
    "Indian FMCG company acquisition deal",
    "India consumer goods merger investment",
    "Indian food company acquisition stake",
    "India personal care brand buyout",
    "Indian startup FMCG acquisition funding",
    "D2C brand India acquisition investment",
    # Private equity / investment
    "FMCG private equity buyout fund",
    "consumer brand venture capital funding",
    "food beverage startup Series funding",
    "consumer goods growth equity investment",
    "FMCG PE backed acquisition deal",
    # Global majors
    "Unilever PepsiCo Nestle acquisition deal",
    "Procter Gamble Colgate acquisition merger",
    "AB InBev Diageo Heineken acquisition",
    "Kraft Heinz Mondelez General Mills deal",
    "LOreal Beiersdorf beauty acquisition",
]
 
FMCG_COMPANIES = [
    # India
    "HUL", "Hindustan Unilever", "Nestle India", "ITC", "Dabur", "Marico",
    "Britannia", "Godrej Consumer", "Godrej Industries", "Colgate Palmolive",
    "Emami", "Patanjali", "Tata Consumer", "Haldirams", "Parle", "Amul",
    "Mother Dairy", "Varun Beverages", "Radico Khaitan", "United Spirits",
    "Bikaji Foods", "Prataap Snacks", "DFM Foods", "Zydus Wellness",
    "Jyothy Labs", "CavinKare", "Wipro Consumer", "Bajaj Consumer",
    "Himalaya", "Nykaa", "Lotus Herbals",
    # Global majors
    "P&G", "Procter Gamble", "Unilever", "Nestle", "PepsiCo", "Coca-Cola",
    "AB InBev", "Mondelez", "Kraft Heinz", "Reckitt", "Church Dwight",
    "Henkel", "LOreal", "Kimberly Clark", "General Mills", "Kellogg",
    "Hershey", "Mars", "Ferrero", "Diageo", "Pernod Ricard", "Edgewell",
    "Energizer", "Clorox", "Beiersdorf", "Essity", "Orkla", "Oatly",
    "Chobani", "Danone", "Lactalis", "Arla Foods", "FrieslandCampina",
    "Associated British Foods", "Tyson Foods", "JBS", "Cargill",
    "Conagra", "Campbell Soup", "Hormel", "McCormick", "Treehouse Foods",
    "Post Holdings", "Lamb Weston", "J M Smucker", "Pilgrim Pride",
    "Spectrum Brands", "Prestige Consumer Healthcare", "Revlon", "Coty",
    "Haleon", "GSK Consumer", "Johnson Johnson Consumer",
    # Beverages
    "Red Bull", "Monster Beverage", "Celsius Holdings", "Vita Coco",
    "Fever Tree", "Nichols", "Britvic", "Tropicana", "Minute Maid",
    "innocent drinks", "Varun Beverages", "Coca Cola FEMSA",
    # Emerging / DTC
    "Beyond Meat", "Impossible Foods", "NotCo", "Oatly", "Mamaearth",
    "Wow Skin Science", "mCaffeine", "Plum", "Sugar Cosmetics",
    "boAt", "Minimalist", "Pilgrim", "Juicy Chemistry",
]
 
FMCG_CATEGORIES = [
    "FMCG", "consumer goods", "food and beverage", "personal care",
    "household products", "packaged goods", "consumer staples",
    "fast moving consumer", "CPG", "consumer packaged", "fmcg sector",
    "beauty brand", "snack brand", "beverage brand", "toiletries",
    "home care", "oral care", "skin care brand", "nutrition brand",
    "health and wellness", "baby care", "feminine hygiene", "pet care",
    "pet food", "hair care", "color cosmetics", "fragrance", "deodorant",
    "detergent brand", "fabric care", "dish care", "air freshener",
    "food company", "dairy brand", "confectionery", "chocolate brand",
    "biscuit brand", "noodles brand", "sauce brand", "condiment brand",
    "spice brand", "edible oil", "instant food", "ready to eat",
    "frozen food", "ice cream brand", "juice brand", "energy drink",
    "sports drink", "water brand", "soft drink", "alcohol brand",
    "beer brand", "wine brand", "spirits brand", "whiskey brand",
    "premium beauty", "drugstore brand", "direct to consumer", "DTC brand",
    "private label", "organic food", "natural products", "plant based",
    "vegan brand", "nutraceutical", "supplement brand", "protein brand",
    "wellness brand", "ayurvedic brand", "herbal brand", "mass market brand",
]
 
DEAL_TERMS = [
    "acquisition", "acquires", "acquired", "merger", "merges", "merged",
    "investment", "invests", "invested", "stake", "buyout", "takeover",
    "deal", "transaction", "M&A", "private equity", "venture capital",
    "funding round", "Series A", "Series B", "Series C", "Series D",
    "joint venture", "JV", "divest", "divestiture", "sells unit",
    "spins off", "spin-off", "carve out", "strategic investment",
    "majority stake", "minority stake", "equity stake", "acqui-hire",
    "leveraged buyout", "LBO", "PE-backed", "management buyout", "MBO",
    "controlling stake", "buys out", "takes over", "merging with",
    "combining with", "strategic alliance", "partnership deal",
    "brand acquisition", "asset purchase", "business combination",
    "tender offer", "bid for", "offer for", "target company",
    "strategic buyer", "financial buyer", "growth equity", "seed round",
    "pre-IPO", "PIPE", "SPAC", "reverse merger", "demerger",
    "restructuring", "consolidation", "roll-up", "bolt-on acquisition",
    "cross-border deal", "hostile takeover", "friendly merger",
    "signed deal", "completed deal", "closed deal", "announced deal",
    "proposed merger", "planned acquisition", "agreed deal", "binding offer",
]
 
 
SOURCE_TIERS = {
    "tier_1": {
        "reuters.com", "ft.com", "bloomberg.com", "wsj.com", "cnbc.com",
        "economictimes.indiatimes.com", "livemint.com", "businessstandard.com",
        "forbes.com", "businessinsider.com", "theguardian.com", "bbc.com",
        "nytimes.com", "washingtonpost.com", "economist.com", "nikkei.com",
    },
    "tier_2": {
        "foodbev.com", "beveragedaily.com", "confectionerynews.com",
        "cosmeticsdesign.com", "retaildetail.eu", "grocerygazette.co.uk",
        "fooddive.com", "cosmeticsbusiness.com", "just-food.com",
        "hbr.org", "mckinsey.com", "pwc.com", "deloitte.com",
        "moneycontrol.com", "ndtv.com", "thehindu.com", "financialexpress.com",
    },
    "tier_3": {
        "yourstory.com", "inc42.com", "techcrunch.com", "vccircle.com",
        "dealstreetasia.com", "businesswire.com", "prnewswire.com",
        "globenewswire.com", "prnewswire.co.uk",
    },
}
 
RSS_FEEDS = [
    ("Reuters Business", "https://feeds.reuters.com/reuters/businessNews"),
    ("Economic Times", "https://economictimes.indiatimes.com/markets/rssfeeds/1977021501.cms"),
    ("Business Standard", "https://www.business-standard.com/rss/home_page_top_stories.rss"),
    ("Livemint", "https://www.livemint.com/rss/companies"),
    ("Financial Express", "https://www.financialexpress.com/feed/"),
]
 
GROQ_MODEL = "qwen/qwen3-32b"

# =============================================================================
# INGESTION
# =============================================================================

def make_id(url: str) -> str:
    return hashlib.sha256(url.encode()).hexdigest()[:16]

def get_domain(url: str) -> str:
    try:
        return urlparse(url).netloc.replace("www.", "")
    except Exception:
        return ""

def fetch_newsapi(api_key: str, days_back: int, status_fn=None) -> list:
    articles = []
    from_date = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d")

    for query in FMCG_QUERIES:
        if status_fn:
            status_fn(f"NewsAPI: querying '{query}'...")
        try:
            resp = requests.get(
                "https://newsapi.org/v2/everything",
                params={
                    "q": query,
                    "from": from_date,
                    "sortBy": "relevancy",
                    "language": "en",
                    "pageSize": 20,
                    "apiKey": api_key,
                },
                timeout=10,
            )
            data = resp.json()
            if data.get("status") != "ok":
                continue
            for item in data.get("articles", []):
                url = item.get("url", "")
                if not url or "removed.com" in url:
                    continue
                articles.append(Article(
                    id=make_id(url),
                    title=(item.get("title") or "").strip(),
                    description=(item.get("description") or "").strip(),
                    url=url,
                    source_name=item.get("source", {}).get("name", ""),
                    source_domain=get_domain(url),
                    published_at=item.get("publishedAt", ""),
                    query_used=query,
                ))
        except Exception:
            pass
        time.sleep(0.25)

    return articles

def fetch_rss(status_fn=None) -> list:
    articles = []
    for feed_name, feed_url in RSS_FEEDS:
        if status_fn:
            status_fn(f"RSS: fetching {feed_name}...")
        try:
            feed = feedparser.parse(feed_url)
            for entry in feed.entries[:25]:
                url = entry.get("link", "")
                if not url:
                    continue
                articles.append(Article(
                    id=make_id(url),
                    title=(entry.get("title") or "").strip(),
                    description=(entry.get("summary") or "").strip(),
                    url=url,
                    source_name=feed_name,
                    source_domain=get_domain(url),
                    published_at=entry.get("published", ""),
                    query_used="rss",
                ))
        except Exception:
            pass
    return articles

def ingest(newsapi_key: str, days_back: int, status_fn=None) -> list:
    articles = []
    if newsapi_key:
        articles.extend(fetch_newsapi(newsapi_key, days_back, status_fn))
    articles.extend(fetch_rss(status_fn))
    return articles

# =============================================================================
# DEDUPLICATION (3 layers, all hosted-safe - no model downloads)
# =============================================================================

def deduplicate(articles: list, fuzzy_thresh: int = 85, tfidf_thresh: float = 0.80) -> tuple:
    stats = {"raw": len(articles), "after_url": 0, "after_fuzzy": 0, "after_tfidf": 0}

    # Layer 1: URL hash (exact dedup)
    seen = set()
    url_unique = []
    for a in articles:
        if a.id not in seen:
            seen.add(a.id)
            url_unique.append(a)
    stats["after_url"] = len(url_unique)

    # Layer 2: Fuzzy title match (rapidfuzz token_sort_ratio)
    fuzzy_unique = []
    for a in url_unique:
        is_dup = any(
            fuzz.token_sort_ratio(a.title, b.title) >= fuzzy_thresh
            for b in fuzzy_unique
        )
        if not is_dup:
            fuzzy_unique.append(a)
    stats["after_fuzzy"] = len(fuzzy_unique)

    # Layer 3: TF-IDF cosine similarity (no model download, pure sklearn)
    # Catches same event reported differently by different outlets
    if len(fuzzy_unique) > 1:
        texts = [f"{a.title} {a.description}" for a in fuzzy_unique]
        try:
            vectorizer = TfidfVectorizer(stop_words="english", max_features=5000, ngram_range=(1, 2))
            matrix = vectorizer.fit_transform(texts)
            sim = cosine_similarity(matrix)
            keep = np.ones(len(fuzzy_unique), dtype=bool)
            for i in range(len(fuzzy_unique)):
                if not keep[i]:
                    continue
                for j in range(i + 1, len(fuzzy_unique)):
                    if sim[i][j] >= tfidf_thresh:
                        keep[j] = False
            final = [a for a, k in zip(fuzzy_unique, keep) if k]
        except Exception:
            final = fuzzy_unique
    else:
        final = fuzzy_unique

    stats["after_tfidf"] = len(final)
    return final, stats

# =============================================================================
# RELEVANCE SCORING
# =============================================================================

def keyword_score(article: Article) -> float:
    text = f"{article.title} {article.description}".lower()
    fmcg_hit = any(t.lower() in text for t in FMCG_COMPANIES + FMCG_CATEGORIES)
    deal_hit = any(t.lower() in text for t in DEAL_TERMS)
    if fmcg_hit and deal_hit:
        return 1.0
    if fmcg_hit or deal_hit:
        return 0.5
    return 0.0

def llm_filter(articles: list, client: Groq, status_fn=None) -> list:
    """LLM relevance check via Groq. Only runs on articles with keyword_score >= 0.5."""
    candidates = [a for a in articles if a.keyword_score >= 0.5]
    if status_fn:
        status_fn(f"LLM filtering {len(candidates)} candidate articles via Groq...")

    for a in candidates:
        try:
            prompt = (
                "You are an FMCG industry analyst. Is this article primarily about M&A activity, "
                "investments, or deal-making in the FMCG/consumer goods industry?\n\n"
                f"Title: {a.title}\nSummary: {a.description[:400]}\n\n"
                "Reply with JSON only (no markdown, no backticks):\n"
                '{"relevant": true or false, "reason": "one sentence", '
                '"deal_type": "acquisition or investment or merger or other or none"}'
            )
            resp = client.chat.completions.create(
                model=GROQ_MODEL,
                max_tokens=150,
                temperature=0.1,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = resp.choices[0].message.content.strip()
            raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()
            raw = re.sub(r"```json|```", "", raw).strip()
            data = json.loads(raw)
            a.llm_relevant = bool(data.get("relevant", False))
            a.deal_type = data.get("deal_type", "other")
            a.llm_reason = data.get("reason", "")
        except Exception:
            # Fallback: trust keyword score if LLM call fails
            a.llm_relevant = a.keyword_score >= 0.5
            a.deal_type = "other"
            a.llm_reason = "LLM fallback - keyword matched"
        time.sleep(1.5)  # Groq is fast, minimal delay needed

    # Mark non-candidates
    for a in articles:
        if a.keyword_score < 0.5:
            a.llm_relevant = False

    return articles

# =============================================================================
# CREDIBILITY SCORING
# =============================================================================

def score_credibility(article: Article) -> tuple:
    domain = article.source_domain
    score_map = {"tier_1": 1.0, "tier_2": 0.7, "tier_3": 0.4}
    for tier, domains in SOURCE_TIERS.items():
        if domain in domains:
            return score_map[tier], tier
    return 0.2, "unverified"

# =============================================================================
# PIPELINE: SCORE + FILTER + RANK
# =============================================================================

def run_scoring(articles: list, client: Groq, status_fn=None) -> list:
    # Step 1: keyword score all
    for a in articles:
        a.keyword_score = keyword_score(a)

    # Step 2: credibility score all
    for a in articles:
        a.credibility_score, a.credibility_tier = score_credibility(a)

    # Step 3: LLM relevance filter (Groq)
    articles = llm_filter(articles, client, status_fn)

    # Step 4: composite score
    for a in articles:
        a.composite_score = round(0.6 * a.keyword_score + 0.4 * a.credibility_score, 3)

    # Step 5: filter + rank
    relevant = [a for a in articles if a.llm_relevant or a.keyword_score == 1.0]
    relevant.sort(key=lambda x: x.composite_score, reverse=True)
    return relevant

# =============================================================================
# NEWSLETTER GENERATION (Groq)
# =============================================================================

def generate_newsletter(articles: list, client: Groq, max_articles: int = 15) -> str:
    top = articles[:max_articles]

    payload = [
        {
            "id": i + 1,
            "title": a.title,
            "description": a.description[:500],
            "source": a.source_name,
            "credibility_tier": a.credibility_tier,
            "deal_type": a.deal_type,
            "composite_score": a.composite_score,
            "published_at": a.published_at,
        }
        for i, a in enumerate(top)
    ]

    prompt = f"""You are a senior FMCG industry analyst and newsletter editor at a top-tier investment research firm. Your job is to write a comprehensive, long-form weekly M&A intelligence report for C-suite executives, PE partners, and investment analysts in the FMCG/consumer goods space.

You have {len(top)} verified and scored news articles from the past week. Write a DETAILED, LONG-FORM newsletter of at least 1500 words covering all angles of FMCG deal activity.

Today's date: {datetime.now().strftime('%B %d, %Y')}

---

OUTPUT FORMAT - use exactly these sections in order:

FMCG M&A Intelligence Report
**Week of {datetime.now().strftime('%B %d, %Y')}**
---

Executive Summary
Write 6-8 crisp bullet points covering the most important developments this week. Each bullet should be a complete, informative sentence that stands alone. Cover deal volumes, key players, and sector themes.

---

Deal of the Week
Pick the single most significant deal or development from the articles. Write a 200-250 word deep-dive: what happened, who the players are, why it matters strategically, what analysts would say about it, and what to watch next. Give it a bold headline.

---

Major Acquisitions
For EACH acquisition article (deal_type="acquisition"):
- **[Company A acquires Company B]** - Write 120-150 words. Cover: deal structure, strategic rationale, what the acquirer gains, competitive implications, deal value if stated, market reaction if mentioned.

---

Investments & Funding
For EACH investment article (deal_type="investment"):
- **[Investor/Fund -> Company]** - Write 120-150 words. Cover: investment thesis, company background, use of funds if mentioned, sector context, why this investment is notable.

---

Mergers & Joint Ventures
For EACH merger/JV article (deal_type="merger"):
- **[Company A + Company B]** - Write 120-150 words. Cover: structure of the deal, combined entity, strategic fit, market position post-merger.

---

Regional Deal Spotlight
Write 150-200 words breaking down deal activity by region this week:
- **India:** What's happening in Indian FMCG M&A
- **North America:** CPG deal trends
- **Europe:** Key developments
- **Asia-Pacific:** Notable activity
Base this only on what's in the articles. If a region has no activity, say so briefly.

---

Strategic Implications
Write 200-250 words of analyst commentary on what this week's deals collectively signal for the FMCG industry. Cover: consolidation trends, which categories are heating up, what acquirers are looking for, valuation environment, private equity activity. This is your editorial/opinion section - write authoritatively.

---

Deals to Watch
List 3-5 situations worth monitoring:
- **[Company/Situation]:** 50-70 words explaining what to watch and why. Tag with [UNVERIFIED SOURCE] if credibility_tier is "unverified".

---

Market Pulse & Outlook
Write 200-250 words on the overall FMCG M&A market tone this week. Cover: deal volume sentiment, which sub-sectors are most active (food, beverage, personal care, household, nutrition), buyer vs seller market dynamics, PE activity levels, what to expect in coming weeks. Be specific and analytical.

---

Source Intelligence Log
Create a table summarizing all articles used:
| # | Headline | Source | Tier | Deal Type | Score |
|---|----------|--------|------|-----------|-------|
(fill in for each article)

---

STRICT RULES - NON-NEGOTIABLE:
1. NEVER hallucinate. Every fact must come from the provided articles.
2. Always name specific companies. Never write "a company" or "an FMCG firm."
3. Include deal values ONLY if explicitly stated in the articles.
4. If a section genuinely has no matching content, write one sentence explaining that and move on. Do NOT skip sections.
5. Tag [UNVERIFIED SOURCE] for any article with credibility_tier="unverified."
6. Write in confident, professional financial analyst voice. No hedging, no filler.
7. Minimum 2500 words. Write every section in full. Use specific numbers, percentages, deal sizes, market share figures, revenue numbers wherever available in the articles. If articles mention company revenues or market positions, include them. Do not pad with vague language - every sentence must carry data or insight.
8. Output clean markdown only. No preamble, no "Here is your newsletter."

Articles JSON:
{json.dumps(payload, indent=2)}"""

    resp = client.chat.completions.create(
        model=GROQ_MODEL,
        max_tokens=6000,
        temperature=0.5,
        messages=[{"role": "user", "content": prompt}],
    )
    
    raw = resp.choices[0].message.content
    clean = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()
    return clean

# =============================================================================
# EXPORTERS
# =============================================================================
def set_tight(para, space_before=0, space_after=2):
    from docx.shared import Pt
    fmt = para.paragraph_format
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)

def to_word(newsletter_md: str) -> bytes:
    if not DOCX_AVAILABLE:
        return b""
    from docx.shared import Pt
 
    doc = Document()
 
    # Global default style - kill spacing on Normal style
    style = doc.styles["Normal"]
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(2)
 
    heading = doc.add_heading("FMCG M&A Intelligence Report", 0)
    heading.alignment = 1
    set_tight(heading, space_before=0, space_after=4)
 
    sub = doc.add_paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M UTC')}")
    set_tight(sub, space_before=0, space_after=6)
 
    for line in newsletter_md.split("\n"):
        line = line.strip()
 
        # Skip blank lines - don't add empty paragraphs (main cause of gaps)
        if not line:
            continue
 
        # Section divider - thin horizontal rule
        elif line == "---":
            p = doc.add_paragraph()
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "4")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "CCCCCC")
            pBdr.append(bottom)
            pPr.append(pBdr)
            set_tight(p, space_before=4, space_after=4)
 
        elif line.startswith("### "):
            h = doc.add_heading(line[4:], level=2)
            set_tight(h, space_before=8, space_after=2)
 
        elif line.startswith("## "):
            h = doc.add_heading(line[3:], level=1)
            set_tight(h, space_before=10, space_after=2)
 
        elif line.startswith("# "):
            h = doc.add_heading(line[2:], level=1)
            set_tight(h, space_before=10, space_after=2)
 
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(line[2:], style="List Bullet")
            set_tight(p, space_before=0, space_after=1)
 
        elif line.startswith("**") and line.endswith("**") and len(line) > 4:
            p = doc.add_paragraph()
            p.add_run(line.strip("*")).bold = True
            set_tight(p, space_before=4, space_after=1)
 
        # Markdown table rows - render as plain text
        elif line.startswith("|"):
            p = doc.add_paragraph(line)
            set_tight(p, space_before=0, space_after=1)
 
        else:
            p = doc.add_paragraph(line)
            set_tight(p, space_before=0, space_after=2)
 
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def to_excel(articles: list, newsletter_md: str) -> bytes:
    if not XLSX_AVAILABLE:
        return b""
    wb = Workbook()

    # Sheet 1: Newsletter
    ws1 = wb.active
    ws1.title = "Newsletter"
    ws1["A1"] = "FMCG M&A Intelligence Report"
    ws1["A1"].font = Font(bold=True, size=16)
    ws1["A2"] = f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}"
    ws1["A2"].font = Font(italic=True)

    row = 4
    for line in newsletter_md.split("\n"):
        cell = ws1.cell(row=row, column=1, value=line)
        if line.startswith("##"):
            cell.font = Font(bold=True, size=13)
            cell.fill = PatternFill("solid", fgColor="E8F4FD")
        elif line.startswith("- ") or line.startswith("* "):
            ws1.cell(row=row, column=1, value="  \u2022 " + line[2:])
        row += 1
    ws1.column_dimensions["A"].width = 120

    # Sheet 2: Scored articles
    ws2 = wb.create_sheet("Scored Articles")
    headers = [
        "Title", "Source", "Domain", "Tier", "Deal Type",
        "Keyword Score", "Credibility Score", "Composite Score",
        "URL", "Published At", "LLM Reason",
    ]
    hfill = PatternFill("solid", fgColor="2E4057")
    hfont = Font(bold=True, color="FFFFFF")

    for col, h in enumerate(headers, 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = hfont
        c.fill = hfill
        c.alignment = Alignment(horizontal="center")

    for ridx, a in enumerate(articles, 2):
        ws2.cell(row=ridx, column=1, value=a.title)
        ws2.cell(row=ridx, column=2, value=a.source_name)
        ws2.cell(row=ridx, column=3, value=a.source_domain)
        ws2.cell(row=ridx, column=4, value=a.credibility_tier)
        ws2.cell(row=ridx, column=5, value=a.deal_type)
        ws2.cell(row=ridx, column=6, value=a.keyword_score)
        ws2.cell(row=ridx, column=7, value=a.credibility_score)
        ws2.cell(row=ridx, column=8, value=a.composite_score)
        ws2.cell(row=ridx, column=9, value=a.url)
        ws2.cell(row=ridx, column=10, value=a.published_at)
        ws2.cell(row=ridx, column=11, value=a.llm_reason)

    col_widths = [60, 20, 25, 12, 15, 15, 18, 17, 50, 25, 40]
    for i, w in enumerate(col_widths, 1):
        ws2.column_dimensions[ws2.cell(row=1, column=i).column_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def to_json(articles: list) -> str:
    return json.dumps([asdict(a) for a in articles], indent=2, default=str)

def to_csv(articles: list) -> str:
    return pd.DataFrame([asdict(a) for a in articles]).to_csv(index=False)

# =============================================================================
# STREAMLIT UI
# =============================================================================

def get_secret(key: str) -> str:
    try:
        return st.secrets[key]
    except Exception:
        return ""

def main():
    st.set_page_config(
        page_title="FMCG M&A NewsAgent",
        page_icon="📰",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── Session state init ────────────────────────────────────────────────────
    if "results" not in st.session_state:
        st.session_state.results = None

    # ── Header ────────────────────────────────────────────────────────────────
    st.title("FMCG M&A Intelligence Newsletter Agent")
    st.divider()

    # ── Sidebar: params only ──────────────────────────────────────────────────
     # ── Sidebar: params only ──────────────────────────────────────────────────
    # Keys read from secrets.toml only - not shown in UI
    groq_key = get_secret("GROQ_API_KEY")
    newsapi_key = get_secret("NEWSAPI_KEY")
 
    with st.sidebar:
        st.header("⚙️ Settings")
 
        days_back = st.slider("Days of news to fetch", 1, 14, 7)
        max_articles = st.slider("Articles in newsletter", 5, 30, 20)
        fuzzy_thresh = st.slider("Fuzzy dedup threshold", 70, 95, 85)
        tfidf_thresh = st.slider("TF-IDF dedup threshold", 0.60, 0.95, 0.80, step=0.01)
 
        st.divider()
 
        run_btn = st.button("🚀 Run Pipeline", type="primary", use_container_width=True)


    # ── Tabs always visible ───────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Pipeline Stats",
        "📰 Newsletter",
        "🗂️ Scored Articles",
        "💾 Downloads",
    ])

    # ── Placeholder state (before first run) ──────────────────────────────────
    if not run_btn and st.session_state.results is None:
        with tab1:
            st.info("Run the pipeline to see stats here.")
        with tab2:
            st.info("Your generated newsletter will appear here.")
        with tab3:
            st.info("Scored and ranked articles will appear here.")
        with tab4:
            st.info("Download options will appear here after pipeline runs.")
        return

    # ── Validate before running ───────────────────────────────────────────────
    if run_btn:
        if not groq_key:
            st.error("Groq API key is required. Get a free one at console.groq.com")
            return

        client = Groq(api_key=groq_key)

        # ── Progress in main area ─────────────────────────────────────────────
        status = st.empty()
        progress = st.progress(0)

        def upd(msg: str):
            status.info(f"⏳ {msg}")

        try:
            upd("Fetching news from NewsAPI + RSS feeds...")
            raw = ingest(newsapi_key, days_back, upd)
            progress.progress(15)

            if not raw:
                status.error("No articles fetched. Check your NewsAPI key or try increasing days.")
                progress.empty()
                return

            upd(f"Deduplicating {len(raw)} articles...")
            deduped, dedup_stats = deduplicate(raw, fuzzy_thresh, tfidf_thresh)
            progress.progress(35)

            upd("Scoring relevance and credibility...")
            scored = run_scoring(deduped, client, upd)
            progress.progress(70)

            if not scored:
                status.warning("No relevant FMCG M&A articles found. Try increasing days.")
                progress.empty()
                return

            upd("Generating newsletter...")
            newsletter = generate_newsletter(scored, client, max_articles)
            progress.progress(100)

            # Store in session state so tabs persist
            st.session_state.results = {
                "raw": raw,
                "scored": scored,
                "dedup_stats": dedup_stats,
                "newsletter": newsletter,
            }

            status.success(
                f"Done! {dedup_stats['raw']} fetched -> "
                f"{dedup_stats['after_tfidf']} after dedup -> "
                f"{len(scored)} relevant articles -> newsletter ready."
            )
            progress.empty()

        except Exception as e:
            err = str(e)
            if "auth" in err.lower() or "api_key" in err.lower() or "invalid" in err.lower():
                status.error("Invalid Groq API key. Check and retry.")
            else:
                status.error(f"Pipeline error: {err}")
            progress.empty()
            with st.expander("Error details"):
                st.exception(e)
            return

    # ── Render results from session state ─────────────────────────────────────
    if st.session_state.results is None:
        return

    raw = st.session_state.results["raw"]
    scored = st.session_state.results["scored"]
    dedup_stats = st.session_state.results["dedup_stats"]
    newsletter = st.session_state.results["newsletter"]

    # Tab 1: Pipeline Stats
    with tab1:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Fetched", dedup_stats["raw"])
        m2.metric("After URL Dedup", dedup_stats["after_url"],
                  delta=f"-{dedup_stats['raw'] - dedup_stats['after_url']}")
        m3.metric("After Fuzzy Dedup", dedup_stats["after_fuzzy"],
                  delta=f"-{dedup_stats['after_url'] - dedup_stats['after_fuzzy']}")
        m4.metric("After TF-IDF Dedup", dedup_stats["after_tfidf"],
                  delta=f"-{dedup_stats['after_fuzzy'] - dedup_stats['after_tfidf']}")
        m5.metric("Relevant Articles", len(scored),
                  delta=f"-{dedup_stats['after_tfidf'] - len(scored)}")

        st.divider()

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Deal Type Breakdown")
            st.bar_chart(pd.Series([a.deal_type for a in scored]).value_counts())
        with col_b:
            st.subheader("Source Credibility Breakdown")
            st.bar_chart(pd.Series([a.credibility_tier for a in scored]).value_counts())

        st.divider()
        st.subheader("Dedup Drop at Each Stage")
        st.dataframe(pd.DataFrame([
            {"Stage": "Raw fetched",          "Articles": dedup_stats["raw"],        "Dropped": 0},
            {"Stage": "After URL hash",        "Articles": dedup_stats["after_url"],  "Dropped": dedup_stats["raw"] - dedup_stats["after_url"]},
            {"Stage": "After fuzzy title",     "Articles": dedup_stats["after_fuzzy"],"Dropped": dedup_stats["after_url"] - dedup_stats["after_fuzzy"]},
            {"Stage": "After TF-IDF",          "Articles": dedup_stats["after_tfidf"],"Dropped": dedup_stats["after_fuzzy"] - dedup_stats["after_tfidf"]},
            {"Stage": "After relevance filter","Articles": len(scored),               "Dropped": dedup_stats["after_tfidf"] - len(scored)},
        ]), use_container_width=True, hide_index=True)

    # Tab 2: Newsletter
    with tab2:
        st.markdown(newsletter)

    # Tab 3: Scored Articles
    with tab3:
        st.caption(f"{len(scored)} relevant FMCG M&A articles found")
        df = pd.DataFrame([{
            "Title": a.title[:90] + ("..." if len(a.title) > 90 else ""),
            "Source": a.source_name,
            "Tier": a.credibility_tier,
            "Deal Type": a.deal_type,
            "Keyword": a.keyword_score,
            "Credibility": a.credibility_score,
            "Composite": a.composite_score,
            "LLM Reason": a.llm_reason,
            "URL": a.url,
        } for a in scored])
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

    # Tab 4: Downloads
    with tab4:
        col1, col2 = st.columns(2)

        with col1:
            if DOCX_AVAILABLE:
                st.download_button(
                    "📄 Newsletter (Word .docx)",
                    data=to_word(newsletter),
                    file_name=f"fmcg_newsletter_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            excel_data = to_excel(scored, newsletter)
            if excel_data:
                st.download_button(
                    "📊 Full Report (Excel .xlsx)",
                    data=excel_data,
                    file_name=f"fmcg_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with col2:
            st.download_button(
                "🗄️ Raw Articles (JSON)",
                data=to_json(raw),
                file_name=f"raw_articles_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True,
            )
            st.download_button(
                "📋 Scored Articles (CSV)",
                data=to_csv(scored),
                file_name=f"scored_articles_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()