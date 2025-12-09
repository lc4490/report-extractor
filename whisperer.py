from unstract.llmwhisperer import LLMWhispererClientV2
from unstract.llmwhisperer.client_v2 import LLMWhispererClientException
import re
import time
import csv
from pathlib import Path
from typing import Dict
import re
import os
import tkinter as tk
from tkinter import filedialog

os.environ["TK_SILENCE_DEPRECATION"] = "1"

# Initialize the client with your API key
# Provide the base URL and API key explicitly
client = LLMWhispererClientV2(base_url="https://llmwhisperer-api.us-central.unstract.com/api/v2", api_key='lBkXnmkVufnr50LkM0hDPGElyGOiQoxIs66e2VFU1_Q')

# HELPER CLEAN FUNCTION
def clean_num(s: str):
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return m.group(0) if m else ""

# BILINGUAL GET LOT NO, WEIGHT, AND THICKNESS
def extract_header_for_page(page_text: str):
    """
    Bilingual header:
    - 訂 單 編 號   24072201-3   重 量 220.0 g/m2   厚 度 0.31 mm
    - Lot no. 25021201-1  Weight 350.5 g/m2  Overall Thickness 0.44 mm
    """
    lot_match = re.search(
        r"(?:訂\s*單\s*編\s*號|Lot no\.)\s+(\S+)",
        page_text
    )

    weight_match = re.search(
        r"(?:重\s*量|Weight)\s+([\d.,]+)",
        page_text
    )

    thickness_match = re.search(
        r"(?:厚\s*度|Overall Thickness).*?([\d]+\.[\d]+)\s*mm",
        page_text,
        flags=re.S,
    )

    if not (lot_match and weight_match and thickness_match):
        return None

    lot_no = lot_match.group(1)
    weight = weight_match.group(1).replace(",", ".")
    thickness = thickness_match.group(1).replace(",", ".")

    return lot_no, weight, thickness

# BILINGUAL FIND SECTION BETWEEN TENSILE AND WHATEVER OTHER VARIABLES, BASICALLY ISOS DATA
def get_tensile_block(page_text: str):
    """
    Find the 'tensile' section:
      - English: 'Item  Tensile Strength(N/5cm)'
      - Chinese: '檢 驗 項 目   拉 力 強 度'
    """
    starts = []

    m_en = re.search(r"Item\s+Tensile Strength", page_text)
    if m_en:
        starts.append(m_en.start())

    m_zh = re.search(r"檢\s*驗\s*項\s*目\s+拉\s*力\s*強\s*度", page_text)
    if m_zh:
        starts.append(m_zh.start())

    if not starts:
        return None

    start = min(starts)
    block = page_text[start:]

    # stop at the next big section (adhesion / high frequency / operator / page break)
    end = len(block)
    for pat in [
        r"\n\s*Item\s+Adhesion Strength",
        r"\n\s*檢\s*驗\s*項\s*目\s+高\s*週\s*波\s*強\s*度",
        r"\n\s*Operator\s*:",
        r"\n\s*檢 人 員",
        r"<<<",
    ]:
        m = re.search(pat, block)
        if m and m.start() < end:
            end = m.start()

    return block[:end]

# BLINGUAL PICK ENGLISH OR CHINESE THEN EXTRACT ROWS
def extract_rows_for_page(page_text: str):
    header = extract_header_for_page(page_text)
    if not header:
        return []

    lot_no, weight, thickness = header

    # If the page has 訂單編號, treat it as Chinese layout
    if re.search(r"訂\s*單\s*編\s*號", page_text):
        print("chinese")
        return extract_hf_rows_for_page_chinese(page_text, lot_no, weight, thickness)

    # Otherwise, assume English UE083-style adhesion block
    print("english")
    return extract_adhesion_rows_english(page_text, lot_no, weight, thickness)

# BILINGUAL EXTRACT ROW FOR ALL PAGES
def extract_all_rows(result_text: str):
    """
    Split by page (<<<) and aggregate rows from each page.
    This automatically handles multiple header sections
    (different weights/thicknesses) and per-page tensile data.
    """
    all_rows = []
    pages = result_text.split("<<<")

    for page_text in pages:
        if not page_text.strip():
            continue
        page_rows = extract_rows_for_page(page_text)
        all_rows.extend(page_rows)

    return all_rows

def collect_results(client: str, path: str):
    rows = []
    try:
        result = client.whisper(
            file_path=path,   
        )
        if result["status_code"] == 202:
            print("Whisper request accepted.")
            print(f"Whisper hash: {result['whisper_hash']}")
            while True:
                print("Polling for whisper status...")
                status = client.whisper_status(whisper_hash=result["whisper_hash"])
                if status["status"] == "processing":
                    print("STATUS: processing...")
                elif status["status"] == "delivered":
                    print("STATUS: Already delivered!")
                    break
                elif status["status"] == "unknown":
                    print("STATUS: unknown...")
                    break
                elif status["status"] == "processed":
                    print("STATUS: processed!")
                    print("Let's retrieve the result of the extraction...")
                    resultx = client.whisper_retrieve(
                        whisper_hash=result["whisper_hash"]
                    )
                    # Refer to documentation for result format
                    result_text = resultx["extraction"]["result_text"]
                    print(result_text)
                    rows = extract_all_rows(result_text)
                    break
                # Poll every 5 seconds
                time.sleep(5)
    except LLMWhispererClientException as e:
        print(e)
    return rows

def big_collect_results(client: str, folder_path: Path, key: str):
    results: list[dict] = []

    for fname in os.listdir(folder_path):
        # 1) Key filter
        if key.lower() not in fname.lower():
            continue

        full_path = folder_path / fname

        # 2) File must exist + be a real file
        if not full_path.is_file():
            continue

        # 3) Must be a PDF file by extension
        if full_path.suffix.lower() != ".pdf":
            continue

        # 4) Process this PDF
        # print(full_path)
        file_rows = collect_results(client, full_path)
        if file_rows:
            results.extend(file_rows)

    return results
# ENGLISH
# ENGNLISH FIND STANDARD TENSILE VALUES
def extract_standard_tensile(page_text: str):
    """
    Extracts the 'Standard' row tensile warp/weft values for the page.
    Pattern is stable across all pages:
    
        Standard
            1800.0   1200.0
    """
    m = re.search(
        r"Standard\s*\n\s*([\d.,]+)\s+([\d.,]+)",
        page_text
    )
    if not m:
        return None, None

    warp = m.group(1).replace(",", ".")
    weft = m.group(2).replace(",", ".")
    return warp, weft

# ENGLISH EXTRACT TENSILE VALUES
def extract_tensile_for_page(page_text: str):
    """
    Optional mapping: per-roll tensile when present (English or Chinese).
    If not present, we just fall back to standard values in row creation.
    """
    block = get_tensile_block(page_text)
    if not block:
        return {}

    tens = {}
    for line in block.splitlines():
        m = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m:
            continue

        roll_no = int(m.group(1))
        rest = m.group(2)
        nums = re.findall(r"[\d]+(?:[.,]\d+)?", rest)
        if len(nums) >= 2:
            warp = nums[0].replace(",", ".")
            weft = nums[1].replace(",", ".")
            tens[roll_no] = (warp, weft)

    return tens

# ENGLISH EXTRACT ALL ROWS
def extract_adhesion_rows_english(page_text: str, lot_no: str, weight: str, thickness: str):
    tensile_map = extract_tensile_for_page(page_text)
    std_warp, std_weft = extract_standard_tensile(page_text)

    rows = []

    adh_match = re.search(
        r"Item\s+Adhesion Strength\(N/5cm\)-B/B(.*?)(?:\n\s*Operator\s*:|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not adh_match:
        return []

    block = adh_match.group(1)

    for line in block.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue

        m = re.match(r"\s*(\d+)\s+", line)
        if not m:
            continue

        roll_no = int(m.group(1))
        rest = line[m.end():]

        parts = rest.split()
        num_tokens = [p for p in parts if re.search(r"\d", p)]

        # Expect at least 8 numeric-ish tokens: B/B(2) + F/B(2) + Tear(2) + Peel(2)
        if len(num_tokens) < 8:
            continue

        nums = [clean_num(p) for p in num_tokens[:8]]

        row = {
            "訂單編號": lot_no,
            "重量": weight,
            "厚度": thickness,
            "roll": roll_no,
            # high-frequency B/B & F/B
            "高週波強度B/B_warp": nums[0],
            "高週波強度B/B_weft": nums[1],
            "高迪波強度F/B_warp": nums[2],
            "高迪波強度F/B_weft": nums[3],
            # tear
            "撕裂強度_warp": nums[4],
            "撕裂強度_weft": nums[5],
            # peel
            "剝離強度_warp": nums[6],
            "剝離強度_weft": nums[7],
        }

        # Tensile: per-roll if present, else standard, else N/A
        if roll_no in tensile_map:
            row["拉力強度_warp"], row["拉力強度_weft"] = tensile_map[roll_no]
        else:
            row["拉力強度_warp"] = std_warp or "N/A"
            row["拉力強度_weft"] = std_weft or "N/A"

        rows.append(row)

    return rows


# CHINESE
# CHINESE FIND ALL STANDARD VALUES
def extract_chinese_standard_mech(page_text: str):
    """
    From the 拉力強度 block, extract standard values:
      tensile_warp, tensile_weft,
      peel_warp, peel_weft,
      tear_warp, tear_weft
    Returns a dict or None.
    """
    m = re.search(
        r"檢\s*驗\s*項\s*目\s+拉\s*力\s*強\s*度(.*?)(?:檢\s*驗\s*項\s*目|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not m:
        return None

    block = m.group(1)

    # jump to 品質標準
    m_std = re.search(r"品\s*質\s*標\s*準(.*)", block, flags=re.S)
    if not m_std:
        return None

    after = m_std.group(1)

    for line in after.splitlines():
        nums = re.findall(r"[\d]+(?:[.,]\d+)?", line)
        # expect 6: tensile_w, tensile_wf, peel_w, peel_wf, tear_w, tear_wf
        if len(nums) >= 6:
            vals = [clean_num(n) for n in nums[:6]]
            return {
                "tensile_warp": vals[0],
                "tensile_weft": vals[1],
                "peel_warp": vals[2],
                "peel_weft": vals[3],
                "tear_warp": vals[4],
                "tear_weft": vals[5],
            }

    return None

# CHINESE EXTRACT ALL VALUES
def extract_chinese_mech_row_map(page_text: str):
    """
    From the 拉力強度 block, build:
      roll_no -> { tensile_warp, tensile_weft, peel_warp, peel_weft, tear_warp, tear_weft }
    For this file, only roll 2 has real values.
    """
    m = re.search(
        r"檢\s*驗\s*項\s*目\s+拉\s*力\s*強\s*度(.*?)(?:檢\s*驗\s*項\s*目\s+高\s*週\s*波\s*強\s*度|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not m:
        return {}

    block = m.group(1)
    rows: Dict[int, Dict[str, str]] = {}

    for line in block.splitlines():
        m_line = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m_line:
            continue

        roll_no = int(m_line.group(1))
        rest = m_line.group(2)

        tokens = rest.split()
        vals = []
        for t in tokens:
            if "ND" in t.upper():
                vals.append("ND")
            elif re.search(r"[\d]", t):
                v = clean_num(t)
                if v:
                    vals.append(v)

        if len(vals) < 6:
            continue

        rows[roll_no] = {
            "tensile_warp": vals[0],
            "tensile_weft": vals[1],
            "peel_warp": vals[2],
            "peel_weft": vals[3],
            "tear_warp": vals[4],
            "tear_weft": vals[5],
        }

    return rows

# CHINESE EXTRACT ALL ROWS
def extract_hf_rows_for_page_chinese(page_text: str, lot_no: str, weight: str, thickness: str):
    """
    Chinese 高週波 parser:
    - Standards come from 拉力強度 block (tensile, peel, tear)
    - Per-roll overrides are applied from 拉力 block (e.g. roll 2)
    """
    rows = []

    # 1. standards
    std = extract_chinese_standard_mech(page_text)
    if std:
        tensile_warp_std = std["tensile_warp"]
        tensile_weft_std = std["tensile_weft"]
        peel_warp_std = std["peel_warp"]
        peel_weft_std = std["peel_weft"]
        tear_warp_std = std["tear_warp"]
        tear_weft_std = std["tear_weft"]
    else:
        tensile_warp_std = tensile_weft_std = "N/A"
        peel_warp_std = peel_weft_std = "N/A"
        tear_warp_std = tear_weft_std = "N/A"

    # 2. per-roll overrides
    mech_rows = extract_chinese_mech_row_map(page_text)

    # 3. now parse 高週波 block
    m = re.search(
        r"檢\s*驗\s*項\s*目\s+高\s*週\s*波\s*強\s*度\s*\(N/in\)-B/B(.*?)(?:檢\s*人\s*員|ISO NO\.|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not m:
        return rows

    block = m.group(1)

    for line in block.splitlines():
        m_line = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m_line:
            continue

        roll_no = int(m_line.group(1))
        rest = m_line.group(2)

        tokens = rest.split()
        nums_raw = [t for t in tokens if re.search(r"\d", t)]
        nums = [clean_num(t) for t in nums_raw if clean_num(t)]

        if len(nums) < 4:
            continue

        bb_warp, bb_weft, fb_warp, fb_weft = nums[:4]

        # start from standards
        tensile_warp = tensile_warp_std
        tensile_weft = tensile_weft_std
        peel_warp = peel_warp_std
        peel_weft = peel_weft_std
        tear_warp = tear_warp_std
        tear_weft = tear_weft_std

        # override if roll has its own row in 拉力 block
        if roll_no in mech_rows:
            info = mech_rows[roll_no]
            tensile_warp = info.get("tensile_warp", tensile_warp)
            tensile_weft = info.get("tensile_weft", tensile_weft)
            peel_warp = info.get("peel_warp", peel_warp)
            peel_weft = info.get("peel_weft", peel_weft)
            tear_warp = info.get("tear_warp", tear_warp)
            tear_weft = info.get("tear_weft", tear_weft)

        row = {
            "訂單編號": lot_no,
            "重量": weight,
            "厚度": thickness,
            "roll": roll_no,
            "拉力強度_warp": tensile_warp,
            "拉力強度_weft": tensile_weft,
            "剝離強度_warp": peel_warp,
            "剝離強度_weft": peel_weft,
            "撕裂強度_warp": tear_warp,
            "撕裂強度_weft": tear_weft,
            "高週波強度B/B_warp": bb_warp,
            "高週波強度B/B_weft": bb_weft,
            "高迪波強度F/B_warp": fb_warp,
            "高迪波強度F/B_weft": fb_weft,
        }

        rows.append(row)

    return rows


# select folder
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select a folder to search")
    root.destroy()
    return folder


def main():
    print("=== Report Extractor ===")
    key = input("Enter search key: ").strip()
    if not key:
        key = ""
    key = ""
    print("Please choose a folder...")
    folder = select_folder()
    if not folder:
        print("No folder selected. Exiting.")
        return

    folder_path = Path(folder)
    rows = big_collect_results(client, folder_path, key)
    # Extract tables from the PDF
    
        
    output_path = Path("output.csv")

    fieldnames = [
        "訂單編號","重量","厚度","roll",
        "拉力強度_warp","拉力強度_weft",
        "剝離強度_warp","剝離強度_weft",
        "撕裂強度_warp","撕裂強度_weft",
        "高週波強度B/B_warp","高週波強度B/B_weft",
        "高迪波強度F/B_warp","高迪波強度F/B_weft",
    ]

    with output_path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
        
if __name__ == "__main__":
    main()