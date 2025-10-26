# geo_dates_from_xml_plus_fields.py
# pip install pandas requests openpyxl beautifulsoup4 urllib3

import time, re, shutil, tempfile
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
import xml.etree.ElementTree as ET

# ========= CONFIG =========
INPUT_FILE = "geo_ids_with_NCBI_Links For all categories reheus.xlsm"  # filename only, in same folder
THROTTLE_SEC = 0.35
LONG_PAUSE_EVERY = 100
LONG_PAUSE_SEC = 2.0
USER_AGENT = "GEO-miniml/1.0 (+contact: your_email@example.com)"
# ==========================

GEO_URL = "https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi"
ACC_RE = re.compile(r"(GSE|GPL|GDS|GSM)\d+", re.I)

BASIC_COLS = [
    "Status", "Submission date", "Last update date",
]
EXTRA_COLS = [
    "Title", "Organism", "Experiment type",
    "Summary", "Overall design",
    "Treatment protocol", "Growth protocol", "Extraction protocol",
]
TECH_COLS = ["Fetch_Status", "Debug_Source", "Error"]

RESULT_COLS = BASIC_COLS + EXTRA_COLS + TECH_COLS

# ---------- HTTP session with retries ----------
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT, "Accept": "*/*", "Connection": "keep-alive"})
    retry = Retry(
        total=6, connect=6, read=6,
        backoff_factor=1.2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s

# ---------- Utilities ----------
def clean_geo_id(raw) -> str:
    raw = "" if pd.isna(raw) else str(raw)
    m = ACC_RE.search(raw.strip())
    return m.group(0).upper() if m else ""

def read_excel_unlocked(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, engine="openpyxl")
    except PermissionError:
        with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as tmp:
            tmp_path = Path(tmp.name)
        shutil.copy2(path, tmp_path)
        try:
            return pd.read_excel(tmp_path, engine="openpyxl")
        finally:
            tmp_path.unlink(missing_ok=True)

def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag

def _norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()

# ---------- XML fetch & parse (namespace-agnostic) ----------
def xml_fetch(session: requests.Session, acc: str) -> bytes:
    params = {"acc": acc, "targ": "self", "form": "xml", "view": "full"}
    r = session.get(GEO_URL, params=params, timeout=30)
    r.raise_for_status()
    return r.content

def xml_extract_fields(xml_bytes: bytes) -> dict:
    """Pull both the 'dates/status' and the additional requested fields from MINiML."""
    empty = {k: "" for k in BASIC_COLS + EXTRA_COLS}
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return empty

    # Find the first top-level entity (Series/Platform/Dataset/Sample).
    entity = None
    for node in root.iter():
        n = _localname(node.tag)
        if n in ("Series", "Platform", "Dataset", "Sample"):
            entity = node
            break
    if entity is None:
        return empty

    out = empty.copy()
    # Collect simple direct-text tags we care about
    tag_map = {
        "Status": "Status",
        "Submission-Date": "Submission date",
        "Last-Update-Date": "Last update date",
        "Title": "Title",
        "Summary": "Summary",
        "Overall-Design": "Overall design",
        "Type": "Experiment type",
        "Treatment-Protocol": "Treatment protocol",
        "Growth-Protocol": "Growth protocol",
        "Extract-Protocol": "Extraction protocol",
        "Organism": "Organism",  # may appear multiple times
    }

    organisms = []
    for child in entity.iter():
        name = _localname(child.tag)
        if name not in tag_map:
            continue
        text = _norm_space(child.text)
        if not text:
            continue
        key = tag_map[name]
        if key == "Organism":
            organisms.append(text)
        else:
            # keep first non-empty, else append if different
            if not out[key]:
                out[key] = text
            elif text not in out[key]:
                out[key] = f"{out[key]}; {text}"

    if organisms:
        # unique + stable order
        seen = []
        for org in organisms:
            if org not in seen:
                seen.append(org)
        out["Organism"] = "; ".join(seen)

    return out

# ---------- HTML fallback for both dt/dd and td/td layouts ----------
def html_fallback(session: requests.Session, acc: str) -> dict:
    params = {"acc": acc}
    r = session.get(GEO_URL, params=params, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    # 1) Build a label->value dictionary from dt/dd (definition list) structure
    dd_map = {}
    for dt in soup.find_all("dt"):
        label = _norm_space(dt.get_text())
        dd = dt.find_next_sibling("dd")
        value = _norm_space(dd.get_text(" ", strip=True)) if dd else ""
        if label:
            dd_map[label.lower()] = value

    # 2) Build a label->value dictionary from table rows (first td = label, second td = value)
    td_map = {}
    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) >= 2:
            label = _norm_space(tds[0].get_text(" ", strip=True)).lower()
            value = _norm_space(tds[1].get_text(" ", strip=True))
            if label:
                td_map[label] = value

    # helper to get either from dd_map or td_map
    def get_label(*names):
        for name in names:
            low = name.lower()
            if low in dd_map and dd_map[low]:
                return dd_map[low]
            if low in td_map and td_map[low]:
                return td_map[low]
        return ""

    out = {k: "" for k in BASIC_COLS + EXTRA_COLS}

    # Basic three
    out["Status"] = get_label("Status")
    out["Submission date"] = get_label("Submission date")
    out["Last update date"] = get_label("Last update date")

    # Extras
    out["Title"] = get_label("Title")
    # Organism can appear as "Organism", "Organisms", or embedded with links.
    org_val = get_label("Organism", "Organisms")
    out["Organism"] = org_val

    out["Experiment type"] = get_label("Experiment type", "Type")
    out["Summary"] = get_label("Summary")
    out["Overall design"] = get_label("Overall design", "Overall Design")

    out["Treatment protocol"] = get_label("Treatment protocol")
    out["Growth protocol"] = get_label("Growth protocol")
    out["Extraction protocol"] = get_label("Extraction protocol", "Extract protocol")

    return out

# ---------- Identify GEO accession column ----------
def get_id_series(df: pd.DataFrame) -> pd.Series:
    for col in df.columns:
        if str(col).strip().lower() == "geo_accession_number":
            return df[col].astype(str)
    return df.iloc[:, 0].astype(str)  # fallback to column A

# ---------- Main ----------
def main():
    in_path = Path(INPUT_FILE)
    if not in_path.exists():
        raise FileNotFoundError(f"Input not found in current folder: {in_path}")

    out_path = in_path.with_name(in_path.stem + "_with_dates.xlsx")
    print(f"Reading: {in_path}")
    df = read_excel_unlocked(in_path)

    # Normalize GEO id column for merge
    ids_raw = get_id_series(df)
    ids_clean = ids_raw.map(clean_geo_id)

    df = df.copy()
    geo_col = None
    for c in df.columns:
        if str(c).strip().lower() == "geo_accession_number":
            geo_col = c; break
    if geo_col is None:
        df["GEO_accession_number"] = ids_clean
    else:
        df.rename(columns={geo_col: "GEO_accession_number"}, inplace=True)
        df["GEO_accession_number"] = ids_clean

    # Drop any pre-existing result columns so our values land in the visible ones
    for c in RESULT_COLS:
        if c in df.columns:
            df.drop(columns=[c], inplace=True)

    sess = make_session()
    rows = []
    successes = 0

    for i, acc in enumerate(df["GEO_accession_number"], start=1):
        if not acc:
            rows.append({"GEO_accession_number": "", **{k: "" for k in BASIC_COLS + EXTRA_COLS},
                         "Fetch_Status": "SKIPPED", "Debug_Source": "", "Error": "No valid GEO ID"})
            continue

        try:
            # XML first
            xml_bytes = xml_fetch(sess, acc)
            fields = xml_extract_fields(xml_bytes)
            source = "XML" if any(fields.values()) else ""

            # If XML produced little/empty, use HTML fallback
            if not any(fields.values()):
                fields = html_fallback(sess, acc)
                source = "HTML" if any(fields.values()) else ""

            ok = any(fields.values())
            rows.append({
                "GEO_accession_number": acc,
                **fields,
                "Fetch_Status": "OK" if ok else "FAILED",
                "Debug_Source": source,
                "Error": "" if ok else "No fields parsed",
            })
            if ok: successes += 1

        except Exception as e:
            # Hard fallback: HTML only
            try:
                fields = html_fallback(sess, acc)
                ok = any(fields.values())
                rows.append({
                    "GEO_accession_number": acc,
                    **fields,
                    "Fetch_Status": "OK" if ok else "FAILED",
                    "Debug_Source": "HTML" if ok else "",
                    "Error": "" if ok else str(e),
                })
                if ok: successes += 1
            except Exception as e2:
                rows.append({
                    "GEO_accession_number": acc,
                    **{k: "" for k in BASIC_COLS + EXTRA_COLS},
                    "Fetch_Status": "FAILED",
                    "Debug_Source": "",
                    "Error": f"{type(e).__name__}: {e} | Fallback: {type(e2).__name__}: {e2}",
                })

        time.sleep(THROTTLE_SEC)
        if i % LONG_PAUSE_EVERY == 0:
            time.sleep(LONG_PAUSE_SEC)

    out_df = pd.DataFrame(rows)

    # Merge and order columns
    final_df = pd.merge(df, out_df, on="GEO_accession_number", how="left")
    ordered = list(df.columns) + [c for c in RESULT_COLS if c not in df.columns]
    final_df = final_df[ordered]

    final_df.to_excel(out_path, index=False)
    print(f"Done: {successes} rows filled. Saved → {out_path}")

if __name__ == "__main__":
    main()
