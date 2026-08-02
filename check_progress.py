#!/usr/bin/env python3
import json, os

state_file = "translation_state.json"
output_file = "eclass_16_ru.xlsx"

if os.path.exists(state_file):
    with open(state_file, "r", encoding="utf-8") as f:
        s = json.load(f)
    cache_count = len(s.get("translation_cache", {}))
    failed_count = len(s.get("failed_translations", {}))
    completed = list(s.get("completed_sheets", {}).keys())
    progress = s.get("sheet_progress", {})
    last_updated = s.get("last_updated", "N/A")
    print(f"State last updated: {last_updated}")
    print(f"Cache entries: {cache_count}")
    print(f"Failed translations: {failed_count}")
    print(f"Completed sheets: {completed}")
    print(f"Progress: {progress}")
else:
    print("No state file yet")

if os.path.exists(output_file):
    print(f"Output file: {os.path.getsize(output_file)} bytes")
else:
    print("No output file yet")
