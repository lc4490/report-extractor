from unstract.llmwhisperer import LLMWhispererClientV2
from unstract.llmwhisperer.client_v2 import LLMWhispererClientException
import re
import time
import csv
from pathlib import Path

# Initialize the client with your API key
# Provide the base URL and API key explicitly
client = LLMWhispererClientV2(base_url="https://llmwhisperer-api.us-central.unstract.com/api/v2", api_key='')
def clean_num(token: str) -> str:
    token = token.strip()
    if token == "-" or token == "":
        return "N/A"
    token = token.replace(",", ".")
    token = token.replace("*", "")
    token = token.replace("'", "")
    return token

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

def extract_header_for_page(page_text: str):
    # Lot no. is stable
    lot_match = re.search(r"Lot no\.\s+(\S+)", page_text)

    # ✅ Allow both "350.5" and "336,0"
    weight_match = re.search(r"Weight\s+([\d.,]+)", page_text)

    # ✅ Thickness: search from "Overall Thickness" up to "mm"
    # and capture the decimal right before "mm" (e.g. 0.44, 0.45).
    # This avoids grabbing the stray "24" on the last page.
    thickness_match = re.search(
        r"Overall Thickness.*?([\d]+\.[\d]+)\s*mm",
        page_text,
        flags=re.S,
    )

    if not (lot_match and weight_match and thickness_match):
        return None

    lot_no = lot_match.group(1)
    weight = weight_match.group(1).replace(",", ".")
    thickness = thickness_match.group(1).replace(",", ".")

    return lot_no, weight, thickness

def extract_tensile_for_page(page_text: str):
    """
    Returns { roll_no: (tensile_warp, tensile_weft) } for this page.
    """
    tensile_map = {}

    block_match = re.search(
        r"Item\s+Tensile Strength.*?Standard(.*?)(?:\n\s*Item\s+Adhesion Strength|\Z)",
        page_text,
        flags=re.S,
    )
    if not block_match:
        return tensile_map  # no tensile block on this page

    block = block_match.group(1)

    for line in block.splitlines():
        m = re.match(r"\s*(\d+)\s+(.*)", line)
        if not m:
            continue

        roll_no = int(m.group(1))
        rest = m.group(2)

        parts = rest.split()
        # remove "Qualified"/"Qualified." etc
        value_tokens = [p for p in parts if not p.startswith("Qualified")]
        num_tokens = [p for p in value_tokens if re.search(r"\d", p)]

        if len(num_tokens) < 2:
            continue

        warp = clean_num(num_tokens[0])
        weft = clean_num(num_tokens[1])

        tensile_map[roll_no] = (warp, weft)

    return tensile_map

def extract_rows_for_page(page_text: str):
    """
    Returns list[dict] of rows for this page only,
    with header info and tensile (if present for that roll).
    """
    header = extract_header_for_page(page_text)
    if not header:
        return []

    lot_no, weight, thickness = header
    tensile_map = extract_tensile_for_page(page_text)
    standard_warp, standard_weft = extract_standard_tensile(page_text)
    rows = []

    # Limit parsing to the Adhesion section → safer
    adh_match = re.search(
        r"Item\s+Adhesion Strength\(N/5cm\)-B/B(.*?)(?:\n\s*Operator\s*:|\Z)",
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

        # remove "Qualified"/"Qualified." tokens
        value_tokens = [p for p in parts if not p.startswith("Qualified")]
        num_tokens = [p for p in value_tokens if re.search(r"\d", p)]

        # Adhesion+Tear+Peel rows have at least 8 numeric-ish tokens
        if len(num_tokens) < 8:
            continue

        nums = [clean_num(p) for p in num_tokens[:8]]

        row = {
            "訂單編號": lot_no,
            "重量": weight,
            "厚度": thickness,
            "roll": roll_no,
            # Adhesion B/B
            "高週波強度B/B_warp": nums[0],
            "高週波強度B/B_weft": nums[1],
            # Adhesion F/B
            "高迪波強度F/B_warp": nums[2],
            "高迪波強度F/B_weft": nums[3],
            # Tear
            "撕裂強度_warp": nums[4],
            "撕裂強度_weft": nums[5],
            # Peel (final column)
            "剝離強度_warp": nums[6],
            "剝離強度_weft": nums[7],
        }

        # Inject tensile if this roll has it on this page
        if roll_no in tensile_map:
            row["拉力強度_warp"], row["拉力強度_weft"] = tensile_map[roll_no]
        else:
            row["拉力強度_warp"] = standard_warp
            row["拉力強度_weft"] = standard_weft

        rows.append(row)

    return rows

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


def main():
    rows = []
    # Extract tables from the PDF
    try:
        result = client.whisper(
            file_path="24072201-3-Magpul-G34-TR.pdf",   
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
                    # rows = extract_all_rows(result_text)
                    break
                # Poll every 5 seconds
                time.sleep(5)
    except LLMWhispererClientException as e:
        print(e)
        
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