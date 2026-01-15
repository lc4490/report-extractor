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
import unicodedata

os.system('cls' if os.name == 'nt' else 'clear')
os.environ["TK_SILENCE_DEPRECATION"] = "1"

# Initialize the client with your API key
# Provide the base URL and API key explicitly
paid_client = LLMWhispererClientV2(base_url="https://llmwhisperer-api.us-central.unstract.com/api/v2", api_key='1s_TV6nE2Q2e3XQTT10ZtXTg8UPLKBjGYs3SWhwbOLk')
client = LLMWhispererClientV2(base_url="https://llmwhisperer-api.us-central.unstract.com/api/v2", api_key='lBkXnmkVufnr50LkM0hDPGElyGOiQoxIs66e2VFU1_Q')

# CLIENT
def is_quota_error(e: Exception):
    msg = str(e).lower()
    return any(k in msg for k in (
        "quota", "limit", "rate", "too many requests", "exceeded",
        "insufficient", "credit", "billing", "payment"
    ))


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
    Bilingual header, OCR-tolerant.
    Returns (lot_no, weight, thickness) or None.
    """
    page_text = re.sub(r"\$\s*(?=\d)", "S", page_text)
    # Lot no / 訂單編號 (allow missing first char due to OCR)
    lot_match = re.search(
        r"(?:.?單\s*編\s*號|[Ll]?ot\s*no\.?)\s*[:：]?\s*([0-9A-Za-z\-]+)",
        page_text
    )

    weight_match = re.search(
        r"(?:重\s*量|Weight)\s*[:：]?\s*([\d.,]+)",
        page_text
    )

    # Thickness: capture number after 厚度 / Overall Thickness
    # Do NOT require "mm" to be immediately after (OCR often separates it)
    thickness_match = re.search(
        r"(?:厚\s*度|Overall\s*Thickness)\s*[:：]?\s*([\d.,]+)",
        page_text,
        flags=re.I
    )

    if not (lot_match and weight_match and thickness_match):
        return None

    lot_no = lot_match.group(1).strip()
    weight = weight_match.group(1).replace(",", ".").strip()
    thickness = thickness_match.group(1).replace(",", ".").strip()

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
    if re.search(r"單\s*編\s*號", page_text):
        # prrint("chinese")
        return extract_hf_rows_for_page_chinese(page_text, lot_no, weight, thickness, filename)
    # Otherwise, assume English UE083-style adhesion block
    # prrint("english")
    return build_rows(page_text, filename)
    # return extract_hf_rows_for_page_english(page_text, lot_no, weight, thickness, filename)

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
        filename_shortened = filename.split("-")
        filename_shortened = filename_shortened[2:]
        filename_shortened = "-".join(filename_shortened)
        filename_shortened = unicodedata.normalize("NFC", str(filename_shortened))
        page_rows = extract_rows_for_page(page_text, filename_shortened)
        all_rows.extend(page_rows)

    return all_rows

def collect_results(client, path: str):
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
        if is_quota_error(e):
            raise  # IMPORTANT: bubble up so we can retry with paid
        print(e)
    return rows

def big_collect_results(client: str, filenames):
    results: list[dict] = []
    
    for full_path in filenames:
        try:
            file_rows = collect_results(client, full_path)
        except LLMWhispererClientException as e:
            if is_quota_error(e):
                file_rows = collect_results(paid_client, full_path)
            else:
                raise            
        if file_rows:
            results.extend(file_rows)
    return results
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

    carry_vals = []
    for line in block.splitlines():
        m_line = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m_line:
            if re.search(r"(合\s*格|判\s*定|\*)", line) and re.search(r"\d", line):
                tokens = line.split()
                vals = []
                for t in tokens:
                    if "ND" in t.upper():
                        vals.append("ND")
                    elif re.search(r"[\d]", t):
                        v = clean_num(t)
                        if v:
                            vals.append(v)
                carry_vals = vals
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

        # if len(vals) < 6:
        #     continue
        if carry_vals:
            vals.extend(carry_vals)
            carry_vals = []
        rows[roll_no] = {
            "tensile_warp": vals[0] if len(vals) > 0 else "N/A",
            "tensile_weft": vals[1] if len(vals) > 1 else "N/A",
            "peel_warp": vals[2] if len(vals) > 2 else "N/A",
            "peel_weft": vals[3] if len(vals) > 3 else "N/A",
            "tear_warp": vals[4] if len(vals) > 4 else "N/A",
            "tear_weft": vals[5] if len(vals) > 5 else "N/A",
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
    r"(?:"
        r"熱\s*(?:壓|封)(?:\s*強\s*度)?"
        r"|高\s*週\s*波\s*強\s*度"
    r")\s*"
    r"\(N\s*/\s*(?:in|2cm)\)\s*-\s*B/B"
    r"(.*?)"
    r"(?="
        r"(?:\s*(?:檢\s*)?驗\s*項\s*目\s*"
        r"(?:"
            r"熱\s*(?:壓|封)(?:\s*強\s*度)?"
            r"|高\s*週\s*波\s*強\s*度"
        r")\s*"
        r"\(N\s*/\s*(?:in|2cm)\)\s*-\s*(?:B/B|F/B))"
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

# ---------- helpers ----------

def clean_num_token(s: str):
    """
    Normalize OCR numeric-ish tokens:
    - strip leading junk like ':' ';'
    - convert comma decimal -> dot
    - keep digits, dot, star, and ND
    """
    s = s.strip()
    if not s:
        return ""
    if re.search(r"\bND\b", s, re.I):
        return "ND"

    # strip leading non-digit chars (keeps negative sign if it exists)
    s = re.sub(r"^[^\d\-]+", "", s)
    s = s.replace(",", ".")
    # keep digits, dot, minus, star
    s = re.sub(r"[^\d\.\-\*]+", "", s)
    # if it ends up empty, return ""
    return s or ""

def first_match(pattern: str, text: str, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None

def split_into_roll_blocks(text: str):
    """
    Split a section that contains roll rows into logical roll blocks.
    A roll block starts with '^\s*\d+\s+' and may continue on subsequent lines.
    """
    blocks = []
    cur = None

    for line in text.splitlines():
        if re.match(r"^\s*\d+\s+", line):
            if cur:
                blocks.append(" ".join(cur))
            cur = [line.strip()]
        elif cur:
            if line.strip():
                cur.append(line.strip())

    if cur:
        blocks.append(" ".join(cur))

    return blocks

def extract_value_decision_pairs(roll_block: str):
    """
    Return list of values that immediately precede Qualified/Unqualified.
    Handles ND too.
    """
    # value token can be ND or number like 31.2* or 120.1 or 75.9-
    pairs = re.findall(
        r"(\bND\b|[^\s]+)\s*(Qualified|Unqualified)",
        roll_block,
        flags=re.I
    )
    vals = []
    for raw, _dec in pairs:
        v = clean_num_token(raw)
        if v:
            # normalize trailing '-' like '75.9-' -> '75.9'
            v = v.rstrip("-")
            vals.append(v)
    return vals

# ---------- core extraction ----------

def extract_page_meta(page_text: str):
    page_text = re.sub(r"\$\s*(?=\d)", "S", page_text)
    lot = first_match(
    r"Lot\s*no\.?\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-_/]*)",
    page_text,
    flags=re.I
) or "N/A"
    weight = first_match(r"Weight\s*[:：]?\s*([0-9]+(?:[.,][0-9]+)?)", page_text, flags=re.I) or "N/A"

    # allow mm to be separated / delayed
    thickness = first_match(
        r"Overall\s*Thickness\s*[:：]?\s*([0-9]+(?:[.,][0-9]+)?)",
        page_text,
        flags=re.I
    ) or "N/A"

    # normalize comma decimals
    weight = weight.replace(",", ".") if weight != "N/A" else weight
    thickness = thickness.replace(",", ".") if thickness != "N/A" else thickness

    return {"lot_no": lot, "weight": weight, "thickness": thickness}

def extract_mech_rows(page_text: str):
    """
    Extract Tensile + Peel from the top mechanical table.
    Very forgiving: grabs roll rows under the top 'Roll no.' section.
    """
    mech_rows: Dict[int, Dict[str, str]] = {}

    # capture from first "Roll no." after the tensile table until the next "Item  Adhesion Strength" or Operator/<<<
    mech_block = first_match(
        r"Roll\s*no\.(.*?)(?=\n\s*Item\s+Adhesion\s+Strength|\n\s*Operator\s*:|<<<|\Z)",
        page_text,
        flags=re.I | re.S
    )
    if not mech_block:
        return mech_rows

    # stitch roll lines and parse value+decision pairs
    for rb in split_into_roll_blocks(mech_block):
        m = re.match(r"^\s*(\d+)\s+(.*)$", rb)
        if not m:
            continue
        roll = int(m.group(1))
        rest = m.group(2)

        vals = extract_value_decision_pairs(rest)
        # Expected order in mech table:
        # tensile_warp, tensile_weft, peel_warp, peel_weft (often peel may be missing)
        tensile_warp = vals[0] if len(vals) > 0 else "N/A"
        tensile_weft = vals[1] if len(vals) > 1 else "N/A"
        peel_warp    = vals[2] if len(vals) > 2 else "N/A"
        peel_weft    = vals[3] if len(vals) > 3 else "N/A"

        mech_rows[roll] = {
            "tensile_warp": tensile_warp,
            "tensile_weft": tensile_weft,
            "peel_warp": peel_warp,
            "peel_weft": peel_weft,
        }

    return mech_rows

def extract_hf_rows(page_text: str):
    """
    Extract Adhesion B/B, Adhesion F/B, Tear, Peel from HF table.
    Uses value+Qualified/Unqualified anchors so wrapping/misalignment is OK.
    """
    hf_rows: Dict[int, Dict[str, str]] = {}

    # capture from HF header until Operator/<<<
    hf_block = first_match(
        r"Item\s+Adhesion\s+Strength\s*\(N/5cm\)\s*-\s*B/B(.*?)(?=\n\s*Operator\s*:|\n\s*ISO\s*NO\.|<<<|\Z)",
        page_text,
        flags=re.I | re.S
    )

    if not hf_block:
        # sometimes OCR drops the "Item" line; fallback: start at "Adhesion Strength...-B/B"
        hf_block = first_match(
            r"Adhesion\s+Strength\s*\(N/5cm\)\s*-\s*B/B(.*?)(?=\n\s*Operator\s*:|\n\s*ISO\s*NO\.|<<<|\Z)",
            page_text,
            flags=re.I | re.S
        )
    if not hf_block:
        return hf_rows

    page_has_fb = re.search(r"\b-\s*F\s*/\s*B\b", page_text, re.I) is not None or re.search(r"\bF\s*/\s*B\b", page_text, re.I) is not None

    # the HF block contains the roll table starting at "Roll no."
    roll_part = first_match(r"Roll\s*no\.?\s*(.*)", hf_block, flags=re.I | re.S)
    if not roll_part:
        # fallback: sometimes "Roll" is broken as "Rol" or "R o l l"
        roll_part = first_match(r"R\s*o\s*l\s*l\s*no\.?\s*(.*)", hf_block, flags=re.I | re.S)
    if not roll_part:
        roll_part = hf_block

    for rb in split_into_roll_blocks(roll_part):
        m = re.match(r"^\s*(\d+)\s+(.*)$", rb)
        if not m:
            continue
        roll = int(m.group(1))
        rest = m.group(2)

        vals = extract_value_decision_pairs(rest)

        # Expected order WITH F/B present:
        # 0 bb_warp, 1 bb_weft, 2 fb_warp, 3 fb_weft, 4 tear_warp, 5 tear_weft, 6 peel_warp, 7 peel_weft
        # Without F/B:
        # 0 bb_warp, 1 bb_weft, 2 tear_warp, 3 tear_weft, 4 peel_warp, 5 peel_weft
        if page_has_fb:
            bb_warp   = vals[0] if len(vals) > 0 else "N/A"
            bb_weft   = vals[1] if len(vals) > 1 else "N/A"
            fb_warp   = vals[2] if len(vals) > 2 else "N/A"
            fb_weft   = vals[3] if len(vals) > 3 else "N/A"
            tear_warp = vals[4] if len(vals) > 4 else "N/A"
            tear_weft = vals[5] if len(vals) > 5 else "N/A"
            peel_warp = vals[6] if len(vals) > 6 else "N/A"
            peel_weft = vals[7] if len(vals) > 7 else "N/A"
        else:
            bb_warp   = vals[0] if len(vals) > 0 else "N/A"
            bb_weft   = vals[1] if len(vals) > 1 else "N/A"
            fb_warp   = "N/A"
            fb_weft   = "N/A"
            tear_warp = vals[2] if len(vals) > 2 else "N/A"
            tear_weft = vals[3] if len(vals) > 3 else "N/A"
            peel_warp = vals[4] if len(vals) > 4 else "N/A"
            peel_weft = vals[5] if len(vals) > 5 else "N/A"

        hf_rows[roll] = {
            "bb_warp": bb_warp, "bb_weft": bb_weft,
            "fb_warp": fb_warp, "fb_weft": fb_weft,
            "tear_warp": tear_warp, "tear_weft": tear_weft,
            "peel_warp": peel_warp, "peel_weft": peel_weft
        }

    return hf_rows

def build_rows(page_text: str, filename: str):
    meta = extract_page_meta(page_text)
    mech = extract_mech_rows(page_text)
    hf   = extract_hf_rows(page_text)

    all_rolls = sorted(set(mech.keys()) | set(hf.keys()))
    rows = []

    for roll in all_rolls:
        m = mech.get(roll, {})
        h = hf.get(roll, {})
        
        peel_warp = m.get("peel_warp", "N/A")
        peel_weft = m.get("peel_weft", "N/A")

        # fallback to HF peel if mech peel missing
        if peel_warp in ("N/A", "", None):
            peel_warp = h.get("peel_warp", "N/A")
        if peel_weft in ("N/A", "", None):
            peel_weft = h.get("peel_weft", "N/A")

        row = {
            "filename": filename,
            "訂單編號": meta["lot_no"],
            "重量": meta["weight"],
            "厚度": meta["thickness"],
            "roll": roll,

            "拉力強度_warp": m.get("tensile_warp", "N/A"),
            "拉力強度_weft": m.get("tensile_weft", "N/A"),

            
            "剝離強度_warp": peel_warp,
            "剝離強度_weft": peel_weft,


            "撕裂強度_warp": h.get("tear_warp", "N/A"),
            "撕裂強度_weft": h.get("tear_weft", "N/A"),

            "高週波強度B/B_warp": h.get("bb_warp", "N/A"),
            "高週波強度B/B_weft": h.get("bb_weft", "N/A"),

            "高迪波強度F/B_warp": h.get("fb_warp", "N/A"),
            "高迪波強度F/B_weft": h.get("fb_weft", "N/A"),
        }
        rows.append(row)

    return rows

def better_value(old: str, new: str) -> str:
    old = old or "N/A"
    new = new or "N/A"

    if old == new:
        return old

    # prefer any real value over N/A
    if old in ("N/A", "") and new not in ("N/A", ""):
        return new
    if new in ("N/A", ""):
        return old

    # prefer numeric over ND
    if old == "ND" and new != "ND":
        return new
    if new == "ND" and old != "ND":
        return old

    # prefer starred values
    if "*" in new and "*" not in old:
        return new
    if "*" in old and "*" not in new:
        return old

    # otherwise keep the old (stable)
    return old

def merge_rows(rows):
    merged = {}
    for r in rows:
        key = (r.get("filename"), r.get("訂單編號"), r.get("roll"))
        if key not in merged:
            merged[key] = r
        else:
            existing = merged[key]
            for k, v in r.items():
                if k in ("roll", "filename", "訂單編號"):
                    continue
                existing[k] = better_value(existing.get(k, "N/A"), v)
    return list(merged.values())


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
    output_path = os.path.join(desktop_path, f"output_{folder_path.name}.csv")

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


    with open(output_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)

        # English header row
        writer.writerow(HEADER_MAP.values())
        merged_rows = merge_rows(rows)
        # Data rows
        for row in merged_rows:
            writer.writerow([row.get(k, "") for k in HEADER_MAP.keys()])

        
if __name__ == "__main__":
    main()