#!/usr/bin/env python3
import json, os

state_file = "translation_state.json"
if os.path.exists(state_file):
    with open(state_file, "r", encoding="utf-8") as f:
        content = f.read()
    cache_start = content.find('"translation_cache"')
    if cache_start >= 0:
        after_key = content.index("{", cache_start)
        failed_start = content.find('"failed_translations"')
        cache_json = content[after_key:failed_start].rsplit("}", 1)[0] + "}"
        cache = json.loads(cache_json)
        print(f"Cache entries: {len(cache)}")
        last_5 = list(cache.items())[-5:]
        for k, v in last_5:
            print(f"  {k[:40]}... -> {v[:40]}...")
