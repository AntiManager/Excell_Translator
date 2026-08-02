#!/usr/bin/env python3
import sys, json
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

import httpx

params = {"client": "dict-chrome-ex", "sl": "auto", "tl": "ru"}
data = {"q": ["Development (Service)", "Concept development", "Mechanical construction"]}

resp = httpx.post("https://translate.googleapis.com/translate_a/t", params=params, data=data, timeout=15.0)
result = resp.json()
for item in result:
    print(item[0])
