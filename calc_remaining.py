#!/usr/bin/env python3
import json, os, sys
from pathlib import Path
sys.path.append(Path(__file__).parent)
import pandas as pd

state_file = "translation_state.json"
if os.path.exists(state_file):
    with open(state_file, encoding="utf-8") as f:
        s = json.load(f)
    cached = len(s.get("translation_cache", {}))
    print(f"Cache entries: {cached}")

file_path = r"C:\Users\evgeniy.bogdanov\Documents\Python\Excell_Translator\eclass_16.xlsx"
df = pd.read_excel(file_path, sheet_name="ECLASS 16.0", dtype=str, keep_default_na=False)

col1 = "Название EN"
col2 = "Определение"
uniq1 = [v for v in df[col1].unique() if v and isinstance(v, str) and len(v.strip()) > 1]
uniq2 = [v for v in df[col2].unique() if v and isinstance(v, str) and len(v.strip()) > 1]
print(f"Unique in '{col1}': {len(uniq1)}")
print(f"Unique in '{col2}': {len(uniq2)}")
print(f"Total unique texts: {len(uniq1) + len(uniq2)}")
print(f"Estimated remaining: {len(uniq1) + len(uniq2) - (cached if os.path.exists(state_file) else 0)}")
print(f"At 5.8 req/sec: {(len(uniq1) + len(uniq2) - (cached if os.path.exists(state_file) else 0)) / 5.8 / 60:.1f} minutes")
