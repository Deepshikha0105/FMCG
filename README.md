# FMCG M&A Intelligence Newsletter Agent

> An automated AI agent that fetches, deduplicates, scores, and summarizes FMCG M&A and investment news into a structured long-form weekly newsletter.

---

## Live Demo

[**Open App on Streamlit Cloud**](https://fmcgdeepshikha.streamlit.app/)

---

## What it does

- Fetches FMCG M&A news from **NewsAPI** (48 targeted queries) + **5 RSS feeds** (Reuters, ET, Business Standard, Livemint, Financial Express)
- Removes duplicates using a **3-layer deduplication pipeline** (URL hash + fuzzy + TF-IDF)
- Scores each article for **FMCG relevance** (keyword scoring + Groq LLM filter) and **source credibility** (tiered whitelist)
- Generates a **detailed long-form newsletter** (1500+ words, 9 sections) via Qwen 3 32B on Groq
- Exports to **Word, Excel, JSON, and CSV**

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    FMCG M&A NewsAgent                           │
│                                                                 │
│  INGESTION            DEDUP                SCORING              │
│  NewsAPI (48Q)  ──>   URL hash      ──>    Keyword score        │
│  5 RSS Feeds          Fuzzy match          LLM filter (Groq)    │
│                       TF-IDF sim           Credibility tier     │
│                                                  │              │
│                                            NEWSLETTER           │
│                                            Qwen 3 32B (Groq)   │
│                                                  │              │
│                                       Word / Excel / JSON / CSV │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │                  Streamlit Demo App                      │   │
│  │  Pipeline Stats | Newsletter | Scored Articles | Downloads│  │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Pipeline Explanation

### Stage 1: Ingestion
- Queries NewsAPI with **48 FMCG-specific search terms** across categories: core M&A, food & beverage, personal care, household, health & nutrition, India-specific, PE/VC, and global majors
- Pulls latest entries from **5 RSS feeds**: Reuters Business, Economic Times, Business Standard, Livemint, Financial Express
- Each article stored as: `id, title, description, url, source_name, source_domain, published_at`

### Stage 2: Deduplication (3 layers)

| Layer | Method | Threshold | Purpose |
|---|---|---|---|
| 1 | URL SHA-256 hash | Exact | Same URL appearing from multiple queries |
| 2 | Fuzzy title match (rapidfuzz) | 85% | "HUL acquires X" vs "HUL acquires X for $50M" |
| 3 | TF-IDF cosine similarity (sklearn) | 0.80 | Same event reported differently by different outlets |

No model downloads needed - TF-IDF runs instantly on CPU via sklearn. Drop count at each layer shown in Pipeline Stats tab.

### Stage 3: Relevance Scoring (hybrid)

**Step A: Keyword Score**
- Checks title + description against 80+ FMCG company names, 55+ category terms, and 60+ deal terms
- Score: 1.0 (both FMCG + deal signal), 0.5 (one signal), 0.0 (no signal)

**Step B: LLM Binary Filter (Groq)**
- Only articles with `keyword_score >= 0.5` sent to LLM (cost control)
- Qwen 3 32B returns: `{ relevant: bool, deal_type: "acquisition/investment/merger/other/none", reason: str }`
- `<think>` block stripped before JSON parse to handle Qwen 3 reasoning output
- 1.5 second delay between calls to avoid Groq rate limits

### Stage 4: Credibility Scoring

Static tiered whitelist - transparent, no black boxes:

| Tier | Score | Examples |
|---|---|---|
| Tier 1 | 1.0 | Reuters, FT, Bloomberg, WSJ, ET, Business Standard, CNBC |
| Tier 2 | 0.7 | FoodBev, FoodDive, Deloitte, HBR, Moneycontrol, Financial Express |
| Tier 3 | 0.4 | VCCircle, BusinessWire, YourStory, PRNewswire, Inc42 |
| Unverified | 0.2 | Anything not in whitelist - flagged as [UNVERIFIED SOURCE] in newsletter |

> **Assumption:** Credibility assessed by source domain reputation only. No real-time fact-checking performed. Whitelist is a static editorial judgment.

### Stage 5: Composite Score + Newsletter

```
composite_score = 0.6 x keyword_score + 0.4 x credibility_score
```

Top N articles by composite score sent to Qwen 3 32B for newsletter generation.

### Newsletter Sections (9 sections, 1500+ words)

1. Executive Summary (6-8 bullets)
2. Deal of the Week (200+ word deep dive)
3. Major Acquisitions (120-150 words each)
4. Investments & Funding (120-150 words each)
5. Mergers & Joint Ventures (120-150 words each)
6. Regional Deal Spotlight (India / North America / Europe / APAC)
7. Strategic Implications (200+ words analyst commentary)
8. Deals to Watch
9. Market Pulse & Outlook (200+ words)
10. Source Intelligence Log (table of all articles used)

---

## Setup

### Prerequisites

- Python 3.10+
- Groq API key - free at [console.groq.com](https://console.groq.com)
- NewsAPI key - free at [newsapi.org](https://newsapi.org) (100 req/day free tier)

### Local Development

```bash
# Clone the repo
git clone https://github.com/CodeWithHarshAI/fmcg-newsletter-agent
cd fmcg-newsletter-agent

# Create virtual environment
py -3.10 -m venv fmcg
fmcg\Scripts\activate        # Windows
# source fmcg/bin/activate   # Mac/Linux

# Install dependencies
pip install -r requirements.txt

# Add API keys
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit secrets.toml with your actual keys

# Run the app
streamlit run app.py
```

### Deploy to Streamlit Cloud (free)

1. Push repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **Create app** -> select repo -> main file: `app.py`
4. Go to **Advanced settings** -> **Secrets** and paste:
   ```toml
   GROQ_API_KEY = "gsk_..."
   NEWSAPI_KEY = "your-newsapi-key"
   ```
5. Click **Deploy** - app live in 2-3 minutes

---

## Output Formats

| Format | Contents |
|---|---|
| Word (.docx) | Full formatted newsletter with tight spacing, section headers, dividers |
| Excel (.xlsx) | Sheet 1: Newsletter content. Sheet 2: All scored articles with all scores |
| JSON | All raw fetched articles with complete metadata |
| CSV | All scored/filtered articles with keyword, credibility, composite scores |

---

## Cost Per Run

| Component | Cost |
|---|---|
| NewsAPI (free tier) | $0 |
| RSS feeds | $0 |
| TF-IDF dedup (local sklearn) | $0 |
| Groq - LLM relevance filter | $0 (free tier) |
| Groq - Newsletter generation | $0 (free tier) |
| **Total** | **$0** |

> Groq free tier is generous enough for full pipeline runs. Upgrade to Dev tier only if hitting TPM limits.

---

## Assumptions and Limitations

1. English-language articles only (`language=en` filter on NewsAPI)
2. No paywall bypass - uses article metadata and description only
3. NewsAPI free tier: 100 requests/day, 1-month historical data max
4. Source credibility is a static whitelist, not a dynamic reputation score
5. LLM does not verify deal values or fact-check between sources
6. RSS feeds do not always provide full article text
7. Groq free tier TPM limit: reduce articles in newsletter slider if hitting 413 errors
8. `<think>` blocks from Qwen 3 reasoning are stripped before parsing

---

## Tech Stack

| Layer | Tool |
|---|---|
| LLM | Qwen 3 32B via Groq API |
| News ingestion | NewsAPI + feedparser (RSS) |
| Dedup L1 | hashlib SHA-256 |
| Dedup L2 | rapidfuzz (fuzzy title match) |
| Dedup L3 | sklearn TF-IDF + cosine similarity |
| Data handling | pandas |
| Word export | python-docx |
| Excel export | openpyxl |
| UI | Streamlit |
| Deployment | Streamlit Community Cloud |

---
