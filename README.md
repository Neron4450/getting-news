# Getting News
**Generate executive‑ready Word reports from current news**: search, scrape, summarize, analyze, and export to a professionally formatted `.docx`—all from one Python file.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-Apache%202.0-green)
![Build](https://github.com/ArcSelf/getting-news/actions/workflows/python.yml/badge.svg)

## Features
- 🔎 **News search** via DuckDuckGo (`ddgs`)
- 🌐 **Web scrape & clean** article bodies (BeautifulSoup4)
- 🧠 **AI summaries & analysis** (OpenAI)
- 📝 **Professional Word document** generation (`python-docx`)
- 📊 **Executive report** with key metrics & insights
- ⚙️ **CLI** with configurable search/scrape/report depth

---

## Requirements
- **Python 3.10+**
- Optional: **OpenAI API key** for AI summaries/analysis

Install dependencies:
```bash
pip install -r requirements.txt
```

**.env (optional)**
```env
# Enables AI summaries/analysis
OPENAI_API_KEY=sk-...
```

---

## Quick Start
```bash
git clone https://github.com/ArcSelf/getting-news.git
cd getting-news

# (optional) create & activate a venv
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate

cp .env.example .env     # then paste your key into .env
pip install -r requirements.txt

python getting_news.py
```

---

## Usage

### Interactive CLI
At launch you’ll be prompted for:
- topic
- articles to search (default 20)
- articles to scrape (default 15)
- number of articles in detailed report (default = scrape count)

### From Python
```python
from getting_news import analyze_news, quick_news_report, comprehensive_news_analysis

print(quick_news_report("federal reserve rate decision"))
print(analyze_news("election security", search_count=25, scrape_count=12, report_detail=10))
print(comprehensive_news_analysis("global energy markets"))
```

---

## Output
A **professional Word document** saved to your working directory, e.g.:
```
Professional_News_Analysis_<topic>_YYYYMMDD_HHMMSS.docx
```
Includes:
- Cover page & metrics
- Table of contents
- Executive dashboard
- Key findings
- Detailed analysis (AI-generated if API key present)
- Article deep‑dive tables
- Statistical analysis
- Source credibility
- Technical appendix (execution log)

---

## Project Structure
```
.
├─ getting_news.py                    # Main program (CLI + engine)
├─ LICENSE                            # Apache-2.0
├─ NOTICE                             # Third-party notices/attribution
├─ README.md
├─ requirements.txt
├─ .env.example
└─ .github/workflows/python.yml       # CI (lint/test)
```

---

## Environment & Privacy
- `OPENAI_API_KEY` is read from `.env` via `python-dotenv`.
- If no key is present, the app still runs; AI summaries/analysis are skipped.
- Scraping should respect site terms and robots.txt. Use responsibly.

---

## Known Limitations
- Some sites block scraping (paywalls/anti-bot).
- Heuristics may miss the main content on atypical layouts.
- Light rate limiting is included; tune if you hit blocks.

---

## License
Licensed under the **Apache License 2.0**. See [`LICENSE`](LICENSE).

## Notices
See [`NOTICE`](NOTICE) for third‑party attributions.

**Direct dependencies:**
- `requests` — Apache 2.0  
- `python-dotenv` — MIT  
- `ddgs` — Apache 2.0  
- `openai` — Apache 2.0  
- `beautifulsoup4` — MIT  
- `python-docx` — MIT  

All other imports are from the Python Standard Library.

---

## Contributing
By submitting a PR, you agree your contributions are licensed under **Apache 2.0** and include the **implicit patent grant** (Section 3).

1. Fork & create a feature branch  
2. Keep functions cohesive and documented  
3. Add/adjust tests or examples if applicable  
4. Update `README.md`/`NOTICE` if you add/modify dependencies
