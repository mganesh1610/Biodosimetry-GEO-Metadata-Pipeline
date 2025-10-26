
# 🧬 Biodosimetry GEO Metadata Pipeline

*Research collaboration support for gene-expression prediction*

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/) [![pandas](https://img.shields.io/badge/pandas-%E2%89%A51.5-150458.svg)](https://pandas.pydata.org/) [![Requests](https://img.shields.io/badge/requests-%E2%89%A52.31-5A29E4.svg)](https://requests.readthedocs.io/) [![BeautifulSoup](https://img.shields.io/badge/bs4-HTML%20parser-6DB33F.svg)](https://www.crummy.com/software/BeautifulSoup/)

> **Goal:** Build a reliable pipeline that mines **NCBI GEO** to compile a clean, analysis-ready metadata table for **radiation-related human/animal studies (GSE/GSM/GPL)** to support **gene-expression prediction**.

---

## 🧭 Table of Contents

* [Overview](#-overview)
* [Key Features](#-key-features)
* [Architecture](#-architecture)
* [What We Extract](#-what-we-extract)
* [End-to-End Workflow](#-end-to-end-workflow)
* [Quickstart](#-quickstart)
* [Scripts](#-scripts)

  * [1) Extract GEO IDs from abstracts](#1-extract-geo-ids-from-abstracts)
  * [2) Harvest GEO metadata (XML → HTML fallback)](#2-harvest-geo-metadata-xml--html-fallback)


---

## 🔎 Overview

This project automates curation of radiation-related studies from **NCBI GEO**. We:

1. Extract **GEO accessions** (GSE/GSM/GPL) from abstracts of radiation studies.
2. Harvest rich metadata from **MINiML XML** (preferred) with **HTML fallback** that handles both `dt/dd` and table layouts.
3. Output a single Excel/CSV that’s **analysis-ready** for downstream **gene-expression prediction**.

---

## ✨ Key Features

* **XML-first, HTML-fallback** parsing with retries/backoff and polite throttling.
* Supports both **definition lists** (`<dt>/<dd>`) and **table layouts** (`<td>label</td><td>value</td>`).
* **Safe Excel I/O** (works even if the source workbook was open).
* Traceability: `Fetch_Status`, `Debug_Source`, `Error`.
* Clean merge that **overwrites stale columns** (no “_x/_y” surprises).

---

## 🏗️ Architecture

```
NCBI abstracts (radiation) ──▶ Regex ID extraction (GSE/GSM/GPL)
                                   │
                                   ├─▶ Seed workbook (IDs + GEO links)
                                   │
                                   └─▶ Harvester (MINiML XML ▶ HTML fallback)
                                            │
                                            └─▶ Final Excel/CSV with rich metadata
```

---

## 🧾 What We Extract

**Dates/Status:** Status · Submission date · Last update date
**Core meta:** Title · Organism (multi; joined with “; ”) · Experiment type
**Design text:** Summary · Overall design
**Protocols:** Treatment protocol · Growth protocol · Extraction protocol

---

## 🔁 End-to-End Workflow

1. **Collect abstracts** for radiation-related studies.
2. **Run regex extractor** to pull GEO IDs (GSE/GSM/GPL).
3. **Create seed workbook**: first column = `GEO_accession_number`, plus optional URL and notes.
4. **Run harvester** to enrich with dates/status, meta, protocols, etc.
5. **Validate** via trace columns; normalize fields as needed; export for modeling.

---

## ⚡ Quickstart

```bash
# clone your repo then:
pip install -r requirements.txt
# or individually:
pip install pandas requests beautifulsoup4 openpyxl urllib3

# 1) Put your abstracts CSV in data/raw/
# 2) Extract GEO IDs:
python scripts/extract_ids_from_abstracts.py

# 3) Seed workbook (optional; or just ensure the xlsm/xlsx exists)
python scripts/make_seed_workbook.py

# 4) Harvest metadata (place next to the workbook)
python scripts/geo_extract_all_fields.py
```

---

## 🧩 Scripts

### 1) Extract GEO IDs from abstracts

`scripts/extract_ids_from_abstracts.py`

```python
import re, pandas as pd, pathlib as p

IN = p.Path("data/raw/radiation_studies_abstracts.csv")
OUT = p.Path("data/processed/geo_ids_from_abstracts.csv")
ID_RE = re.compile(r"\b(GSE|GSM|GPL|GDS)\d+\b", re.I)

df = pd.read_csv(IN)
rows = []
for _, r in df.iterrows():
    text = " ".join(str(r.get(c, "")) for c in ["title", "abstract"])
    ids = sorted(set(m.group(0).upper() for m in ID_RE.finditer(text)))
    for acc in ids:
        rows.append({"source_record": r.get("id"), "GEO_accession_number": acc})

pd.DataFrame(rows).drop_duplicates().to_csv(OUT, index=False)
print("Saved:", OUT)
```

### 2) Harvest GEO metadata (XML → HTML fallback)

`/scripts/geo_extract_all_fields.py` (trimmed core — full script in repo)

```python
# pip install pandas requests beautifulsoup4 openpyxl urllib3
import re, time, shutil, tempfile, pandas as pd
from pathlib import Path
import requests, xml.etree.ElementTree as ET
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup

INPUT_FILE = "geo_ids_with_NCBI_Links For all categories.xlsm"
GEO_URL = "https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi"
ACC_RE  = re.compile(r"(GSE|GPL|GDS|GSM)\d+", re.I)
BASIC   = ["Status","Submission date","Last update date"]
EXTRA   = ["Title","Organism","Experiment type","Summary","Overall design",
           "Treatment protocol","Growth protocol","Extraction protocol"]
RESULTS = BASIC + EXTRA + ["Fetch_Status","Debug_Source","Error"]

def session():
    s = requests.Session()
    retry = Retry(total=6, backoff_factor=1.2,
                  status_forcelist=[429,500,502,503,504],
                  allowed_methods={"GET"})
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.headers.update({"User-Agent":"GEO-miniml/1.0 (+contact: you@example.com)"})
    return s

def norm(s): return re.sub(r"\s+"," ", (s or "")).strip()
def local(t): return t.rsplit("}",1)[-1] if "}" in t else t

def read_xlsx(path: Path):
    try: return pd.read_excel(path, engine="openpyxl")
    except PermissionError:
        with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as tmp:
            tmp_path = Path(tmp.name)
        shutil.copy2(path, tmp_path)
        try: return pd.read_excel(tmp_path, engine="openpyxl")
        finally: tmp_path.unlink(missing_ok=True)

def xml_fetch(sess, acc):
    r = sess.get(GEO_URL, params={"acc":acc,"targ":"self","form":"xml","view":"full"}, timeout=30)
    r.raise_for_status()
    return r.content

def xml_extract(xml_bytes):
    out = {k:"" for k in BASIC+EXTRA}
    try: root = ET.fromstring(xml_bytes)
    except ET.ParseError: return out
    entity = next((n for n in root.iter() if local(n.tag) in ("Series","Platform","Dataset","Sample")), None)
    if entity is None: return out
    tag_map = {"Status":"Status","Submission-Date":"Submission date","Last-Update-Date":"Last update date",
               "Title":"Title","Summary":"Summary","Overall-Design":"Overall design","Type":"Experiment type",
               "Treatment-Protocol":"Treatment protocol","Growth-Protocol":"Growth protocol",
               "Extract-Protocol":"Extraction protocol","Organism":"Organism"}
    orgs=[]
    for c in entity.iter():
        name, txt = local(c.tag), norm(c.text)
        if name not in tag_map or not txt: continue
        key = tag_map[name]
        if key=="Organism": orgs.append(txt)
        else: out[key] = txt if not out[key] else out[key] if txt in out[key] else f"{out[key]}; {txt}"
    if orgs: out["Organism"]="; ".join(dict.fromkeys(orgs))
    return out

def html_extract(sess, acc):
    r = sess.get(GEO_URL, params={"acc":acc}, timeout=30); r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    dd, td = {}, {}
    for dt in soup.find_all("dt"):
        lab = norm(dt.get_text()); ddn = dt.find_next_sibling("dd")
        dd[lab.lower()] = norm(ddn.get_text(" ", strip=True)) if ddn else ""
    tables = soup.find_all("table", attrs={"width": re.compile(r"^\s*600\s*$")}) or soup.find_all("table")
    for tbl in tables:
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds)>=2:
                lab = norm(tds[0].get_text(" ", strip=True)).lower()
                val = norm(tds[1].get_text(" ", strip=True))
                td[lab] = val
    def get(*names):
        for n in names:
            k=n.lower()
            if dd.get(k): return dd[k]
            if td.get(k): return td[k]
        return ""
    out = {k:"" for k in BASIC+EXTRA}
    out["Status"]            = get("Status")
    out["Submission date"]   = get("Submission date")
    out["Last update date"]  = get("Last update date")
    out["Title"]             = get("Title")
    out["Organism"]          = get("Organism","Organisms")
    out["Experiment type"]   = get("Experiment type","Type","Sample type")
    out["Summary"]           = get("Summary")
    out["Overall design"]    = get("Overall design","Overall Design")
    out["Treatment protocol"]= get("Treatment protocol","Treatment Protocol")
    out["Growth protocol"]   = get("Growth protocol","Growth Protocol")
    out["Extraction protocol"]=get("Extraction protocol","Extract protocol","Extraction Protocol","Extract Protocol")
    return out
```

*(Full script in repo includes merging logic, throttling, and Excel writer.)*

---

