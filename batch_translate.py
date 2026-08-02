#!/usr/bin/env python3
import sys, time, json, os, re
from pathlib import Path
sys.path.append(Path(__file__).parent)

import pandas as pd
import httpx

STATE_FILE = "translation_state.json"
FILE_PATH = r"C:\Users\evgeniy.bogdanov\Documents\Python\Excell_Translator\eclass_16.xlsx"
OUTPUT_PATH = r"C:\Users\evgeniy.bogdanov\Documents\Python\Excell_Translator\eclass_16_ru.xlsx"
COLUMNS = ["Название EN", "Определение"]
BATCH_SIZE = 50
MAX_RETRIES = 3

def should_translate(text):
    if not isinstance(text, str):
        return False
    text = text.strip()
    if not text or len(text) < 2:
        return False
    if text.replace(".", "").replace(",", "").isdigit():
        return False
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        return False
    alpha = sum(1 for c in text if c.isalpha())
    if alpha / max(len(text), 1) < 0.3:
        return False
    return True

def load_cache():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            s = json.load(f)
        return s.get("translation_cache", {})
    return {}

def save_cache(cache):
    state = {"translation_cache": cache, "last_updated": time.time()}
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def translate_batch(texts):
    params = {"client": "dict-chrome-ex", "sl": "auto", "tl": "ru"}
    data = {"q": texts}

    for attempt in range(MAX_RETRIES):
        try:
            resp = httpx.post(
                "https://translate.googleapis.com/translate_a/t",
                params=params,
                data=data,
                timeout=30.0
            )
            resp.raise_for_status()
            result = resp.json()
            translations = [part[0] for part in result]
            return translations
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                wait = 2 ** attempt
                print(f"  retry {attempt+1}: {e} (wait {wait}s)")
                time.sleep(wait)
            else:
                print(f"  FAILED batch: {e}")
                return list(texts)

print("Reading Excel...")
df = pd.read_excel(FILE_PATH, sheet_name="ECLASS 16.0", dtype=str, keep_default_na=False)
print(f"Rows: {len(df)}, Cols: {len(df.columns)}")

all_unique = {}
for col in COLUMNS:
    vals = [v for v in df[col].unique().tolist() if should_translate(v)]
    all_unique[col] = vals
    print(f"  {col}: {len(vals)} unique translatable")

cache = load_cache()
print(f"Cache before: {len(cache)} entries")

uncached = []
seen = set()
for col in COLUMNS:
    for t in all_unique[col]:
        if t not in cache and t not in seen:
            uncached.append(t)
            seen.add(t)

print(f"Need to translate: {len(uncached)}")

if not uncached:
    print("Nothing to translate.")
    for col in COLUMNS:
        df[col] = df[col].map(cache).fillna(df[col])
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ECLASS 16.0", index=False)
    print(f"Saved: {OUTPUT_PATH}")
    sys.exit(0)

start = time.time()
done = 0

for i in range(0, len(uncached), BATCH_SIZE):
    batch = uncached[i:i + BATCH_SIZE]
    t0 = time.time()
    results = translate_batch(batch)
    for orig, trans in zip(batch, results):
        cache[orig] = trans
    done += len(batch)
    elapsed = time.time() - t0
    total_elapsed = time.time() - start
    rate = done / total_elapsed if total_elapsed > 0 else 0
    eta = (len(uncached) - done) / rate if rate > 0 else 0
    print(f"[{done}/{len(uncached)}] batch {len(batch)} in {elapsed:.1f}s | {rate:.1f}/sec | ETA {eta/60:.1f}min")
    if (i // BATCH_SIZE) % 5 == 0:
        save_cache(cache)

save_cache(cache)
print(f"\nApplying translations...")
for col in COLUMNS:
    df[col] = df[col].map(cache).fillna(df[col])

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="ECLASS 16.0", index=False)

total = time.time() - start
print(f"Done! {total:.0f}s ({total/60:.1f}min)")
print(f"Translated: {done}, Cache: {len(cache)}")
print(f"Output: {OUTPUT_PATH}")
