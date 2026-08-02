#!/usr/bin/env python3
import sys
from pathlib import Path
import pandas as pd
sys.path.append(Path(__file__).parent)
from translator import ExcelTranslator

file_path = r"C:\Users\evgeniy.bogdanov\Documents\Python\Excell_Translator\eclass_16.xlsx"

translator = ExcelTranslator()

df = pd.read_excel(file_path, sheet_name="ECLASS 16.0", dtype=str, keep_default_na=False)
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print(f"Column dtypes: {dict(df.dtypes)}")
print()

col = "Название EN"
print(f"--- Column: {col} ---")
unique = df[col].unique().tolist()
print(f"Unique count: {len(unique)}")
for v in unique[:10]:
    should = translator._should_translate(v)
    alpha = sum(1 for c in v if c.isalpha()) / len(v) if v and len(v) > 0 else 0
    print(f"  [{should}] alpha_ratio={alpha:.2f} len={len(v)} repr={repr(v[:60])}")

print()

col2 = "Определение"
print(f"--- Column: {col2} ---")
unique2 = df[col2].unique().tolist()
print(f"Unique count: {len(unique2)}")
for v in unique2[:10]:
    should = translator._should_translate(v)
    alpha = sum(1 for c in v if c.isalpha()) / len(v) if v and len(v) > 0 else 0
    print(f"  [{should}] alpha_ratio={alpha:.2f} len={len(v)} repr={repr(v[:60])}")

print()
print("--- estimate_sheet_volume ---")
vol = translator.estimate_sheet_volume(file_path, "ECLASS 16.0", [col, col2])
print(f"Volume: {vol}")
