# FMCG M&A Intelligence Newsletter Agent

> An automated agent that fetches, deduplicates, scores, and summarizes FMCG M&A and investment news into a structured weekly newsletter.

---

## Live Demo

[**Open App on Streamlit Cloud**](https://your-app-name.streamlit.app) *(replace after deploy)*

---

## What it does

- Fetches FMCG M&A news from **NewsAPI** + **5 RSS feeds** (Reuters, ET, Business Standard, Livemint, Financial Express)
- Removes duplicates using a **3-layer deduplication pipeline**
- Scores each article for **FMCG relevance** (keyword + LLM) and **source credibility** (tiered whitelist)
- Generates a **structured newsletter** via Claude (Anthropic)
- Exports to **Word, Excel, JSON, and CSV**

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    FMCG M&A NewsAgent                           │
│                                                                 │
│  INGESTION          DEDUP              SCORING                  │
│  NewsAPI     ──>    URL hash   ──>     Keyword score            │
│  RSS Feeds          Fuzzy match        LLM filter (Claude)      │
│                     Semantic sim       Credibility tier         │
│                                              │                  │
│                                         NEWSLETTER              │
│                                         Claude claude-sonnet-4-6          │
│                                              │                  │
│                                    Word / Excel / JSON / CSV    │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │              Streamlit Demo App                          │   │
│  │  Pipeline Stats | Newsletter | Scored Articles | Downloads│  │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Pipeline Explanation

### Stage 1: Ingestion
- Queries NewsAPI with 8 FMCG-specific search terms (e.g. `"FMCG acquisition merger"`, `"consumer goods deal investment"`)
- Pulls latest entries from 5 RSS feeds
- Each article stored as: `id, title, description, url, source_name, source_domain, published_at`

### Stage 2: Deduplication (3 layers)

| Layer | Method | Threshold | Purpose |
|---|---|---|---|
| 1 | URL SHA-256 hash | Exact | Same URL from multiple queries |
| 2 | Fuzzy title match (rapidfuzz) | 85% | "HUL acquires X" vs "HUL acquires X for $50M" |
| 3 | Semantic similarity (sentence-transformers) | 0.88 cosine | Same event, different headlines from different outlets |

Layer 3 uses `all-MiniLM-L6-v2` running locally on CPU. The drop count at each layer is displayed in the app's Pipeline Stats tab.

### Stage 3: Relevance Scoring (hybrid)

**Step A: Keyword Score**
- Checks title + description against a list of 30+ FMCG company names, 20+ category terms, and 25+ deal terms
- Score: 1.0 (both FMCG + deal terms hit), 0.5 (one category hit), 0.0 (no hit)

**Step B: LLM Binary Filter**
- Only articles with `keyword_score >= 0.5` are sent to Claude (cost control)
- Claude returns: `{ relevant: bool, deal_type: "acquisition/investment/merger/other/none", reason: str }`
- Cost: ~$0.02 per pipeline run (30 articles x 200 tokens each)

### Stage 4: Credibility Scoring

Static tiered whitelist (transparent, no black boxes):

| Tier | Score | Examples |
|---|---|---|
| Tier 1 | 1.0 | Reuters, FT, Bloomberg, WSJ, ET, Business Standard |
| Tier 2 | 0.7 | FoodBev, Deloitte, HBR, Moneycontrol, Financial Express |
| Tier 3 | 0.4 | VCCircle, BusinessWire, YourStory, PRNewswire |
| Unverified | 0.2 | Anything not in the whitelist (flagged in newsletter) |

> **Assumption:** Credibility is assessed by source domain reputation only. No real-time fact-checking is performed. The whitelist is a static editorial judgment.

### Stage 5: Composite Score + Newsletter

```
composite_score = 0.6 * keyword_score + 0.4 * credibility_score
```

Top N articles by composite score are sent to Claude for newsletter generation. The newsletter prompt strictly instructs Claude not to hallucinate, to flag unverified sources, and to include deal values only when explicitly stated.

---

## Setup

### Prerequisites

- Python 3.11+
- Anthropic API key ([console.anthropic.com](https://console.anthropic.com))
- NewsAPI key ([newsapi.org](https://newsapi.org)) - free tier, 100 req/day

### Local Development

```bash
# Clone the repo
git clone https://github.com/your-username/fmcg-newsletter-agent
cd fmcg-newsletter-agent

# Install dependencies
pip install -r requirements.txt

# Add API keys
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit secrets.toml with your actual keys

# Run the app
streamlit run app.py
```

### Deploy to Streamlit Cloud (free)

1. Push this repo to GitHub (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** -> select your repo -> set main file to `app.py`
4. Go to **Advanced settings** -> **Secrets** and add:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   NEWSAPI_KEY = "your-newsapi-key"
   ```
5. Click **Deploy**

> If you don't want to commit API keys: leave `secrets.toml` out of the repo. The app has text input fields in the sidebar so users can enter keys manually.

---

## Output Formats

| Format | Contents |
|---|---|
| Word (.docx) | Formatted newsletter with all sections |
| Excel (.xlsx) | Sheet 1: Newsletter content. Sheet 2: All scored articles with scores |
| JSON | All raw fetched articles with full metadata |
| CSV | All scored/filtered articles with relevance + credibility scores |

---

## Cost Per Run

| Component | Cost |
|---|---|
| NewsAPI (free tier) | $0 |
| RSS feeds | $0 |
| sentence-transformers (local CPU) | $0 |
| Claude - LLM relevance filter (~30 x 200 tokens) | ~$0.02 |
| Claude - Newsletter generation (~3000 tokens) | ~$0.01 |
| **Total** | **~$0.03** |

---

## Assumptions and Limitations

1. English-language articles only (`language=en` filter on NewsAPI)
2. No paywall bypass - uses article metadata and description only
3. NewsAPI free tier limited to 100 requests/day and 1-month historical data
4. Source credibility is a static whitelist, not a dynamic reputation score
5. LLM does not verify deal values or check for contradictions between sources
6. RSS feeds do not always provide full article text
7. Semantic dedup requires sentence-transformers; falls back to 2-layer dedup if unavailable

---

## Tech Stack

- **LLM**: Claude claude-sonnet-4-6 (Anthropic)
- **News**: NewsAPI + feedparser (RSS)
- **Dedup**: rapidfuzz + sentence-transformers (all-MiniLM-L6-v2)
- **Scoring**: keyword rules + Claude binary filter
- **Exports**: python-docx, openpyxl
- **UI**: Streamlit
- **Deployment**: Streamlit Community Cloud

---

## Author

Built by [Your Name] | [GitHub](https://github.com/your-username) | [LinkedIn](https://linkedin.com/in/yourprofile)
