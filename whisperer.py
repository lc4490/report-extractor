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
import sys
import ctypes

os.system('cls' if os.name == 'nt' else 'clear')
os.environ["TK_SILENCE_DEPRECATION"] = "1"

# Initialize the client with your API key
# Provide the base URL and API key explicitly
client = LLMWhispererClientV2(base_url="https://llmwhisperer-api.us-central.unstract.com/api/v2", api_key='1s_TV6nE2Q2e3XQTT10ZtXTg8UPLKBjGYs3SWhwbOLk')

# HELPER CLEAN FUNCTION
def clean_num(s: str):
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return m.group(0) if m else ""

def is_num_less_than_2(x):
    try:
        return float(x) < 2
    except:
        return False      # not a number

def get_desktop_path():
    if sys.platform.startswith("win"):
        CSIDL_DESKTOPDIRECTORY = 0x10
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_DESKTOPDIRECTORY, None, 0, buf
        )
        return buf.value
    else:
        return os.path.join(os.path.expanduser("~"), "Desktop")

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
def extract_rows_for_page(page_text: str, filename):
    header = extract_header_for_page(page_text)
    if not header:
        return []

    lot_no, weight, thickness = header

    # If the page has 訂單編號, treat it as Chinese layout
    if re.search(r"訂\s*單\s*編\s*號", page_text):
        # prrint("chinese")
        return extract_hf_rows_for_page_chinese(page_text, lot_no, weight, thickness, filename)

    # Otherwise, assume English UE083-style adhesion block
    # prrint("english")
    return extract_hf_rows_for_page_english(page_text, lot_no, weight, thickness, filename)

# BILINGUAL EXTRACT ROW FOR ALL PAGES
def extract_all_rows(result_text: str, filename):
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
        page_rows = extract_rows_for_page(page_text, filename)
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
                    rows = extract_all_rows(result_text, path.name)
                    break
                # Poll every 5 seconds
                time.sleep(5)
    except LLMWhispererClientException as e:
        print(e)
    return rows

def big_collect_results(client: str, filenames):
    results: list[dict] = []
    
    for full_path in filenames:
        print(full_path.name)
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
def extract_hf_rows_for_page_english(page_text: str, lot_no: str, weight: str, thickness: str, filename):
    """
    Chinese 高週波 / 熱壓 parser (schema-preserving):

    - Tensile / peel / tear standards from 拉力強度 block
    - Per-roll overrides from 拉力 block (if present)
    - 熱壓 or 高週波 B/B & F/B from this block
      * IMPORTANT: Never infer F/B unless an explicit "-F/B" header exists on the page.
      * If F/B header is absent, F/B stays "N/A" and is NOT filled from standards.

    NOTE: Output dict keys are kept EXACTLY the same as your old function
          (including "高迪波強度F/B_*" typo) to match existing CSV headers.
    """
    rows = []

    # 1) Mechanical standards (tensile / peel / tear)
    std_mech = extract_chinese_standard_mech(page_text)
    if std_mech:
        tensile_warp_std = std_mech["tensile_warp"]
        tensile_weft_std = std_mech["tensile_weft"]
        peel_warp_std    = std_mech["peel_warp"]
        peel_weft_std    = std_mech["peel_weft"]
        tear_warp_std    = std_mech["tear_warp"]
        tear_weft_std    = std_mech["tear_weft"]
    else:
        tensile_warp_std = tensile_weft_std = "N/A"
        peel_warp_std    = peel_weft_std    = "N/A"
        tear_warp_std    = tear_weft_std    = "N/A"

    # 2) HF standards (B/B & F/B)
    std_hf = extract_chinese_hf_standard(page_text)
    if std_hf:
        bb_warp_std = std_hf.get("bb_warp", "N/A")
        bb_weft_std = std_hf.get("bb_weft", "N/A")
        fb_warp_std = std_hf.get("fb_warp", "N/A")
        fb_weft_std = std_hf.get("fb_weft", "N/A")
    else:
        bb_warp_std = bb_weft_std = "N/A"
        fb_warp_std = fb_weft_std = "N/A"

    # 3) Per-roll mechanical overrides
    mech_rows = extract_english_mech_row_map(page_text)

    # 4) Only treat F/B as real if an explicit header exists anywhere on the page
    page_has_fb_header = re.search(r"\bF\s*/\s*B\b", page_text) is not None
    if not page_has_fb_header:
        fb_warp_std = fb_weft_std = "N/A"

    # 5) Extract ALL B/B blocks (handles side-by-side duplicated tables)
    #    Supports both 熱壓 and 高週波強度
    bb_block_pat = re.compile(
        r"(?:Item\s*)"
        r"(?:Adhesion\s*Strength)\s*"
        r"\(N/5cm\)\s*-\s*B/B"
        r"(.*?)"
        r"(?="
        r"(?:\s*Item\s*(?:Adhesion\s*Strength)\s*\(N/5cm\)\s*-\s*(?:B/B|F/B))"
        r"|(?:\s*Operator\s*:)"
        r"|(?:\s*ISO\s*NO\.)"
        r"|(?:\s*<<<)"
        r"|\Z"
        r")",
        flags=re.S | re.I
    )

    bb_blocks = [m.group(1) for m in bb_block_pat.finditer(page_text)]
    if not bb_blocks:
        return rows

    block = "\n".join(b.strip() for b in bb_blocks if b and b.strip())


    peel = re.search(
        r"P\s*e\s*e\s*l\s*[\.\-\/]?\s*S\s*t\s*r\s*e\s*n\s*g\s*t\s*h",
        block,
        re.IGNORECASE | re.DOTALL
    )
    tear = re.search(
        r"T\s*e\s*a\s*r\s*[\.\-\/]?\s*S\s*t\s*r\s*e\s*n\s*g\s*t\s*h",
        block,
        re.IGNORECASE | re.DOTALL
    )

    for line in block.splitlines():
        if not re.search(r"(Qualified|Unqualified)", line, flags=re.I):
            continue

        m_line = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m_line:
            continue

        roll_no = int(m_line.group(1))
        rest = m_line.group(2)

        # ---- collect numeric/ND tokens on this row ----
        tokens = []
        for tok in rest.split():
            if "ND" in tok.upper():
                tokens.append("ND")
            elif re.search(r"\d", tok):
                num = clean_num(tok)
                if num and not is_num_less_than_2(num):
                    tokens.append(num)

        if not tokens:
            continue

        # defaults
        bb_warp = "N/A"
        bb_weft = "N/A"
        fb_warp = "N/A"
        fb_weft = "N/A"
    
            # ===== CASE 2: Normal layout (no peel column) =====
            # Never infer F/B unless the page has an explicit F/B header.
        print(tokens)
        if page_has_fb_header:
            if len(tokens) >= 4:
                bb_warp, bb_weft, fb_warp, fb_weft = tokens[:4]
            elif len(tokens) == 2:
                bb_warp, bb_weft = tokens
        else:
            # Only B/B is valid; ignore any extra numbers (often duplicated B/B table or other cols)
            if len(tokens) >= 2:
                bb_warp, bb_weft = tokens[:2]
        if tear and peel and len(tokens) >= 8:
            if peel.start() < tear.start():
                peel_warp, peel_weft, tear_warp, tear_weft = tokens[4:]
            else:
                tear_warp,tear_weft, peel_warp, peel_weft = tokens[4:]
        elif tear and len(tokens)>=6:
            tear_warp, tear_weft = tokens[4:]
        elif peel and len(tokens)>=6:
            peel_warp, peel_weft = tokens[4:]
            
        # ---- fill missing HF from standards ----
        if bb_warp == "N/A" and bb_warp_std != "N/A":
            bb_warp = bb_warp_std
        if bb_weft == "N/A" and bb_weft_std != "N/A":
            bb_weft = bb_weft_std

        # Only fill F/B from standards if page truly has F/B
        if page_has_fb_header:
            if fb_warp == "N/A" and fb_warp_std != "N/A":
                fb_warp = fb_warp_std
            if fb_weft == "N/A" and fb_weft_std != "N/A":
                fb_weft = fb_weft_std

        # ---- tensile / peel / tear from standards + mech overrides ----
        tensile_warp = tensile_warp_std
        tensile_weft = tensile_weft_std
        if not peel: 
            peel_warp    = peel_warp_std
            peel_weft    = peel_weft_std
        if not tear:
            tear_warp    = tear_warp_std
            tear_weft    = tear_weft_std

        if roll_no in mech_rows:
            info = mech_rows[roll_no]
            tensile_warp = info.get("tensile_warp", tensile_warp)
            tensile_weft = info.get("tensile_weft", tensile_weft)
            peel_warp    = info.get("peel_warp", peel_warp)
            peel_weft    = info.get("peel_weft", peel_weft)
            tear_warp    = info.get("tear_warp", tear_warp)
            tear_weft    = info.get("tear_weft", tear_weft)

        # Only override peel_warp when the HF block truly has a peel column

        # IMPORTANT: keep output keys EXACTLY as before (incl. your "高迪波..." typo)
        row = {
            "filename": filename,
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

def extract_english_mech_row_map(page_text: str):
    """
    From the 拉力強度 block, build:
      roll_no -> { tensile_warp, tensile_weft, peel_warp, peel_weft, tear_warp, tear_weft }
    For this file, only roll 2 has real values.
    """
    m = re.search(
        r"(?:Item\s*)?Tensile\s*Strength\s*\(N/5cm\)"
        r"(.*?)(?="
        r"(?:\s*Item\s*(?:Adhesion\s*Strength)\s*\(N/5cm\)\s*-\s*B/B)"  # next section (HF)
        r"|(?:\s*Operator\s*:)"
        r"|(?:\s*ISO\s*NO\.)"
        r"|(?:\s*<<<)"
        r"|\Z"
        r")",
        page_text,
        flags=re.S | re.I
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
            if len(vals) >= 2:
                rows[roll_no] = {
                    "tensile_warp": vals[0],
                    "tensile_weft": vals[1],
                }
            return rows

        rows[roll_no] = {
            "tensile_warp": vals[0],
            "tensile_weft": vals[1],
            "peel_warp": vals[2],
            "peel_weft": vals[3],
            "tear_warp": vals[4],
            "tear_weft": vals[5],
        }

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
    return None
    m = re.search(
        r"(?:檢)?\s*驗\s*項\s*目\s+拉\s*力\s*強\s*度(.*?)(?:檢\s*驗\s*項\s*目|<<<|\Z)",
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
    # prrint("returned end")
    return None

def extract_chinese_hf_standard(page_text: str):
    """
    Extract standard values for:
      - B/B_warp
      - B/B_weft
      - F/B_warp
      - F/B_weft

    Works for both:
      - 熱壓 (N/in)-B/B
      - 高週波強度 (N/in)-B/B

    Will NOT reuse B/B for F/B.
    If F/B standards are missing, they remain "N/A".
    """
    return None

    # Optional but helpful: only consider F/B standards if an explicit F/B header exists somewhere
    page_has_fb_header = re.search(r"\(N/in\)\s*-\s*F/B", page_text) is not None

    # Match the HF section starting at the B/B header (either 熱壓 or 高週波強度)
    m = re.search(
        r"(?:檢\s*)?驗\s*項\s*目\s*"
        r"(?:熱\s*壓|高\s*週\s*波\s*強\s*度)\s*"
        r"\(N/in\)\s*-\s*B/B"
        r"(.*?)(?:(?:[:：]?\s*(?:檢\s*)?驗\s*人\s*員\s*[:：])|ISO\s*NO\.|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not m:
        return None

    block = m.group(1)

    # Jump to 品質標準 section
    m_std = re.search(r"品\s*質\s*標\s*準(.*)", block, flags=re.S)
    if not m_std:
        return None

    after = m_std.group(1)

    bb_warp = "N/A"
    bb_weft = "N/A"
    fb_warp = "N/A"
    fb_weft = "N/A"

    # Find the first numeric line after 品質標準
    for line in after.splitlines():
        # grab numbers like 200.0 or 180,0
        nums = re.findall(r"[\d]+(?:[.,]\d+)?", line)
        if not nums:
            continue

        # B/B always first two numbers
        if len(nums) >= 1:
            bb_warp = clean_num(nums[0])
        if len(nums) >= 2:
            bb_weft = clean_num(nums[1])

        # F/B only if:
        #  - the page actually has an F/B header somewhere, AND
        #  - we actually see 3rd/4th numbers here
        if page_has_fb_header and len(nums) >= 3:
            fb_warp = clean_num(nums[2])
        if page_has_fb_header and len(nums) >= 4:
            fb_weft = clean_num(nums[3])

        break

    return {
        "bb_warp": bb_warp,
        "bb_weft": bb_weft,
        "fb_warp": fb_warp,
        "fb_weft": fb_weft,
    }

    """
    Extract standard values for:
      - B/B_warp
      - B/B_weft
      - F/B_warp
      - F/B_weft

    Will NOT reuse B/B for F/B. If F/B standards are missing, they remain "N/A".
    """
    m = re.search(
        r"(?:檢)?\s*驗\s*項\s*目\s+熱\s*壓\s*\(N/in\)-B/B(.*?)(?:(?:檢\s*)?驗\s*人\s*員|[:：]\s*驗\s*人\s*員\s*[:：]|ISO NO\.|<<<|\Z)",
        page_text,
        flags=re.S,
    )
    if not m:
        return None

    block = m.group(1)

    # Jump to 品質標準 section
    m_std = re.search(r"品\s*質\s*標\s*準(.*)", block, flags=re.S)
    if not m_std:
        return None

    after = m_std.group(1)

    bb_warp = "N/A"
    bb_weft = "N/A"
    fb_warp = "N/A"
    fb_weft = "N/A"

    # Find numeric lines immediately after 品質標準
    for line in after.splitlines():
        nums = re.findall(r"[\d]+(?:[.,]\d+)?", line)
        if not nums:
            continue

        # B/B always comes first
        if len(nums) >= 1:
            bb_warp = clean_num(nums[0])
        if len(nums) >= 2:
            bb_weft = clean_num(nums[1])

        # F/B ONLY if explicitly present
        if len(nums) >= 3:
            fb_warp = clean_num(nums[2])
        if len(nums) >= 4:
            fb_weft = clean_num(nums[3])

        break  # Only process the first numeric line after 標準

    return {
        "bb_warp": bb_warp,
        "bb_weft": bb_weft,
        "fb_warp": fb_warp,
        "fb_weft": fb_weft,
    }

# CHINESE EXTRACT ALL VALUES
def extract_chinese_mech_row_map(page_text: str):
    """
    From the 拉力強度 block, build:
      roll_no -> { tensile_warp, tensile_weft, peel_warp, peel_weft, tear_warp, tear_weft }
    For this file, only roll 2 has real values.
    """
    m = re.search(
        r"(?:檢\s*)?驗\s*項\s*目\s*拉\s*力\s*強\s*度\s*\(N/in\)"
        r"(.*?)(?="
        r"(?:\s*(?:檢\s*)?驗\s*項\s*目\s*(?:高\s*週\s*波\s*強\s*度|熱\s*壓)\s*\(N/in\)\s*-\s*B/B)"  # next section (HF)
        r"|(?:\s*[:：]?\s*(?:檢\s*)?驗\s*人\s*員\s*[:：])"
        r"|(?:\s*ISO\s*NO\.)"
        r"|(?:\s*<<<)"
        r"|\Z"
        r")",
        page_text,
        flags=re.S
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
def extract_hf_rows_for_page_chinese(page_text: str, lot_no: str, weight: str, thickness: str, filename):
    """
    Chinese 高週波 / 熱壓 parser (schema-preserving):

    - Tensile / peel / tear standards from 拉力強度 block
    - Per-roll overrides from 拉力 block (if present)
    - 熱壓 or 高週波 B/B & F/B from this block
      * IMPORTANT: Never infer F/B unless an explicit "-F/B" header exists on the page.
      * If F/B header is absent, F/B stays "N/A" and is NOT filled from standards.

    NOTE: Output dict keys are kept EXACTLY the same as your old function
          (including "高迪波強度F/B_*" typo) to match existing CSV headers.
    """
    rows = []

    # 1) Mechanical standards (tensile / peel / tear)
    std_mech = extract_chinese_standard_mech(page_text)
    if std_mech:
        tensile_warp_std = std_mech["tensile_warp"]
        tensile_weft_std = std_mech["tensile_weft"]
        peel_warp_std    = std_mech["peel_warp"]
        peel_weft_std    = std_mech["peel_weft"]
        tear_warp_std    = std_mech["tear_warp"]
        tear_weft_std    = std_mech["tear_weft"]
    else:
        tensile_warp_std = tensile_weft_std = "N/A"
        peel_warp_std    = peel_weft_std    = "N/A"
        tear_warp_std    = tear_weft_std    = "N/A"

    # 2) HF standards (B/B & F/B)
    std_hf = extract_chinese_hf_standard(page_text)
    if std_hf:
        bb_warp_std = std_hf.get("bb_warp", "N/A")
        bb_weft_std = std_hf.get("bb_weft", "N/A")
        fb_warp_std = std_hf.get("fb_warp", "N/A")
        fb_weft_std = std_hf.get("fb_weft", "N/A")
    else:
        bb_warp_std = bb_weft_std = "N/A"
        fb_warp_std = fb_weft_std = "N/A"

    # 3) Per-roll mechanical overrides
    mech_rows = extract_chinese_mech_row_map(page_text)

    # 4) Only treat F/B as real if an explicit header exists anywhere on the page
    page_has_fb_header = re.search(r"\(N/in\)\s*-\s*F/B", page_text) is not None
    if not page_has_fb_header:
        fb_warp_std = fb_weft_std = "N/A"

    # 5) Extract ALL B/B blocks (handles side-by-side duplicated tables)
    #    Supports both 熱壓 and 高週波強度
    bb_block_pat = re.compile(
        r"(?:檢\s*)?驗\s*項\s*目\s*"
        r"(?:熱\s*壓|高\s*週\s*波\s*強\s*度)\s*"
        r"\(N/in\)\s*-\s*B/B"
        r"(.*?)"
        r"(?="
        r"(?:\s*(?:檢\s*)?驗\s*項\s*目\s*(?:熱\s*壓|高\s*週\s*波\s*強\s*度)\s*\(N/in\)\s*-\s*(?:B/B|F/B))"
        r"|(?:\s*[:：]?\s*(?:檢\s*)?驗\s*人\s*員\s*[:：])"
        r"|(?:\s*ISO\s*NO\.)"
        r"|(?:\s*<<<)"
        r"|\Z"
        r")",
        flags=re.S
    )
    bb_blocks = [m.group(1) for m in bb_block_pat.finditer(page_text)]
    if not bb_blocks:
        return rows

    block = "\n".join(b.strip() for b in bb_blocks if b and b.strip())

    hf_has_peel = ("剝 離 強 度" in block) or ("剝離 強 度" in block) or ("剝離強度" in block)

    for line in block.splitlines():
        if not re.search(r"(合\s*格|不\s*合\s*格)", line):
            continue

        m_line = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m_line:
            continue

        roll_no = int(m_line.group(1))
        rest = m_line.group(2)

        # ---- collect numeric/ND tokens on this row ----
        tokens = []
        for tok in rest.split():
            if "ND" in tok.upper():
                tokens.append("ND")
            elif re.search(r"\d", tok):
                num = clean_num(tok)
                if num and not is_num_less_than_2(num):
                    tokens.append(num)

        if not tokens:
            continue

        # defaults
        bb_warp = "N/A"
        bb_weft = "N/A"
        fb_warp = "N/A"
        fb_weft = "N/A"
        peel_extra_warp = None

        if hf_has_peel:
            # ===== CASE 1: HF block has 剝離強度 column (legacy weird layout) =====
            if len(tokens) >= 5:
                # B/B(2) + F/B(2) + 剝離_warp
                bb_warp, bb_weft, fb_warp, fb_weft = tokens[:4]
                peel_extra_warp = tokens[4]

            elif len(tokens) == 4:
                bb_warp, bb_weft, fb_warp, fb_weft = tokens

            elif len(tokens) == 3:
                bb_warp, bb_weft, third = tokens

                if third == "ND":
                    peel_extra_warp = third
                else:
                    treated_as_peel = False
                    try:
                        v = float(third)
                        peel_ref = float(peel_warp_std) if peel_warp_std not in ("N/A", "ND") else None
                        fb_ref   = float(fb_warp_std) if fb_warp_std not in ("N/A", "ND") else None
                        if peel_ref is not None and fb_ref is not None:
                            if abs(v - peel_ref) <= abs(v - fb_ref) * 0.7:
                                treated_as_peel = True
                        elif peel_ref is not None and v <= peel_ref * 2:
                            treated_as_peel = True
                    except ValueError:
                        pass

                    if treated_as_peel:
                        peel_extra_warp = third
                    else:
                        if page_has_fb_header:
                            fb_warp = third

            elif len(tokens) == 2:
                bb_warp, bb_weft = tokens

        else:
            # ===== CASE 2: Normal layout (no peel column) =====
            # Never infer F/B unless the page has an explicit F/B header.
            if page_has_fb_header:
                if len(tokens) >= 4:
                    bb_warp, bb_weft, fb_warp, fb_weft = tokens[:4]
                elif len(tokens) == 3:
                    bb_warp, bb_weft, fb_warp = tokens
                elif len(tokens) == 2:
                    bb_warp, bb_weft = tokens
            else:
                # Only B/B is valid; ignore any extra numbers (often duplicated B/B table or other cols)
                if len(tokens) >= 2:
                    bb_warp, bb_weft = tokens[:2]

        # ---- fill missing HF from standards ----
        if bb_warp == "N/A" and bb_warp_std != "N/A":
            bb_warp = bb_warp_std
        if bb_weft == "N/A" and bb_weft_std != "N/A":
            bb_weft = bb_weft_std

        # Only fill F/B from standards if page truly has F/B
        if page_has_fb_header:
            if fb_warp == "N/A" and fb_warp_std != "N/A":
                fb_warp = fb_warp_std
            if fb_weft == "N/A" and fb_weft_std != "N/A":
                fb_weft = fb_weft_std

        # ---- tensile / peel / tear from standards + mech overrides ----
        tensile_warp = tensile_warp_std
        tensile_weft = tensile_weft_std
        peel_warp    = peel_warp_std
        peel_weft    = peel_weft_std
        tear_warp    = tear_warp_std
        tear_weft    = tear_weft_std

        if roll_no in mech_rows:
            info = mech_rows[roll_no]
            tensile_warp = info.get("tensile_warp", tensile_warp)
            tensile_weft = info.get("tensile_weft", tensile_weft)
            peel_warp    = info.get("peel_warp", peel_warp)
            peel_weft    = info.get("peel_weft", peel_weft)
            tear_warp    = info.get("tear_warp", tear_warp)
            tear_weft    = info.get("tear_weft", tear_weft)

        # Only override peel_warp when the HF block truly has a peel column
        if hf_has_peel and peel_extra_warp is not None:
            peel_warp = peel_extra_warp

        # IMPORTANT: keep output keys EXACTLY as before (incl. your "高迪波..." typo)
        row = {
            "filename": filename,
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
    folder = input("貼上資料夾路徑（或將資料夾拖曳到此處）：").strip().strip('"')
    p = Path(folder)
    if not p.exists() or not p.is_dir():
        print("無效的資料夾。")
        return None
    return p
    # root = tk.Tk()
    # root.withdraw()
    # root.attributes("-topmost", True)
    # folder = filedialog.askdirectory(title="Select a folder to search")
    # root.destroy()
    # return folder
    
# parse folders
def parse_only_pdfs(folder_path):
    folder_path = Path(folder_path)
    filenames = []

    for path in folder_path.rglob("*.pdf"):
        # 1) Must be a file (rglob can return dirs in other patterns)
        if not path.is_file():
            continue

        # 2) Only keep files with "TR" in the filename
        if "TR" not in path.name:
            continue

        filenames.append(path)

    return filenames
    
# enter filters
def enter_filters(filenames):
    original = filenames[:]   # shallow copy
    current = filenames[:]

    while True:
        print("\n目前檔案：")
        for fname in current:
            print(fname.name)

        key = input(
            "\n輸入搜尋字詞（輸入 * 或 all 重置，直接按回車鍵完成）："
        ).strip().lower()

        if not key:
            break

        if key in ("*", "all"):
            current = original[:]
            continue

        current = [
            fname for fname in current
            if key in fname.name.lower()
        ]

    return current


def main():
    print("=== Report Extractor ===")
    print("請選擇資料夾...")
    folder = select_folder()
    if not folder:
        print("沒有資料夾")
        return
    
    folder_path = Path(folder)
    
    filenames = parse_only_pdfs(folder_path)
    
    filenames = enter_filters(filenames)
            
    rows = big_collect_results(client, filenames)
    # Extract tables from the PDF
    desktop_path = get_desktop_path()
    output_path = os.path.join(desktop_path, "output.csv")

    HEADER_MAP = {
        "filename": "Filename",
        "訂單編號": "Lot No.",
        "重量": "Weight",
        "厚度": "Overall Thickness",
        "roll": "roll",

        "拉力強度_warp": "Tensile_Strength_warp",
        "拉力強度_weft": "Tensile_Strength_weft",

        "剝離強度_warp": "Peel_Strength_warp",
        "剝離強度_weft": "Peel_Strength_weft",

        "撕裂強度_warp": "Tear_Strength_warp",
        "撕裂強度_weft": "Tear_Strength_weft",

        "高週波強度B/B_warp": "Adhesion_Strength_B/B_warp",
        "高週波強度B/B_weft": "Adhesion_Strength_B/B_weft",

        "高迪波強度F/B_warp": "Adhesion_Strength_F/B_warp",
        "高迪波強度F/B_weft": "Adhesion_Strength_F/B_weft",
    }


    with open(output_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # English header row
        writer.writerow(HEADER_MAP.values())

        # Data rows
        for row in rows:
            writer.writerow([row.get(k, "") for k in HEADER_MAP.keys()])

        
if __name__ == "__main__":
    main()