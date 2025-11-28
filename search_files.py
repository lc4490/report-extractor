import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import platform
from pdf2image import convert_from_path
from openai import OpenAI
import base64
from io import BytesIO
import json
import csv

os.system('cls' if os.name == 'nt' else 'clear')
os.environ["TK_SILENCE_DEPRECATION"] = "1"

# converts image to base64
def pil_to_b64(img):
    """Convert PIL image to base64 PNG."""
    buf = BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")

# runs ocr on a pdf with a path
def ocr_pdf_with_openai(pdf_path: Path, client):
    """
    Convert first page of PDF to image, send to OpenAI Vision,
    and return ONLY numbers (including -, /, .).
    """
    pages = convert_from_path(
        str(pdf_path),
        dpi=350,              
        fmt="png",            
        thread_count=1       
    )

    # Take first page for now
    page = pages[0]
    img_b64 = pil_to_b64(page)

    system_prompt = """你是一個精準的資料擷取引擎。 你會收到一張紡織/布料測試報告的圖片。你的工作是從圖片中的「檢驗結果」欄位中，擷取指定欄位的**第一筆資料（支號 1）**，並將結果以「dictionary（字典）」格式輸出。 ⚠️ 通用規則： - 只能讀取「檢驗結果」，不能使用「標準」或其他欄位。 - 不得輸出任何單位（例如 g/m2, mm, N/in, N）。 - 如果欄位存在而且有數值，就一定要用該數值，不能寫成 "N/A"。 - 如果欄位顯示 ND 或 N/A，請如實輸出（例如 "ND"）。 - 如果「在整張報告中仔細查找後」，確定該欄位完全不存在或該格完全沒有任何數字/文字，才輸出 "N/A"。 - key 名稱必須完全符合下列指定名稱。 - value 一律為字串格式。 - 只能輸出一個 JSON dictionary，不得包含說明文字、註解或額外內容。 📌 需輸出的欄位（全部都必須給出一個值，如果找不到就用 "N/A"）： - 訂單編號 - 重量 - 厚度 - 拉力強度_warp - 拉力強度_weft - 剝離強度_warp - 剝離強度_weft - 撕裂強度_warp - 撕裂強度_weft - 高週波強度B/B_warp - 高週波強度B/B_weft - 高迪波強度F/B_warp - 高迪波強度F/B_weft 📌 關於 F/B 欄位的**特別規則**（非常重要）： 1. 先在整張報告中**仔細尋找**「高週波強度 (N/in)-F/B」或類似標題，以及對應的 Warp / Weft「檢驗結果」。 2. 如果能找到 F/B 的欄位，且在「檢驗結果」中有數字或文字，就一定要輸出該值： - 例如："高迪波強度F/B_warp": "111.5", "高迪波強度F/B_weft": "87.5"。 - 這種情況**絕對不能**輸出 "N/A"。 3. 只有在以下情況，才可以輸出 "N/A"： - 整張報告中沒有出現任何 F/B 的標題或欄位（完全沒有 F/B 區塊），或 - 有 F/B 區塊，但該格「檢驗結果」完全空白、看不到任何數字或文字。 4. 如果只有其中一個方向缺值（例如 Warp 有值、Weft 沒有），那就： - 有值的方向 → 輸出實際數值； - 沒值的方向 → 輸出 "N/A"。 📌 輸出格式範例（僅為示意）： { "訂單編號": "24072201-3", "重量": "220.0", "厚度": "0.31", "拉力強度_warp": "974.8", "拉力強度_weft": "518.9", "剝離強度_warp": "ND", "剝離強度_weft": "ND", "撕裂強度_warp": "26.7", "撕裂強度_weft": "41.2", "高週波強度B/B_warp": "215.5", "高週波強度B/B_weft": "187.4", "高迪波強度F/B_warp": "111.5", "高迪波強度F/B_weft": "87.5" } 只能輸出上述結構的字典，不得輸出任何其他內容。"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        temperature=0,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{img_b64}",
                        },
                    }
                ],
            },
        ],
    )

    content = response.choices[0].message.content

    try:
        return json.loads(content)
    except json.JSONDecodeError:
        print("Model did not return valid JSON. Raw content below:\n")
        print(content)
        return []

# closes terminal when the program ends
def close_terminal_if_frozen():
    # if program not frozen do not kill program
    if not getattr(sys, 'frozen', False):
        return  

    # finds out which system the user is using, closes terminal based on that
    system = platform.system()

    if system == "Windows":
        os.system(f"taskkill /F /PID {os.getpid()}")

    elif system == "Darwin":
        os.system("osascript -e 'tell application \"Terminal\" to close first window' || true")

    elif system == "Linux":
        # Optionally do nothing or close terminal emulator
        pass

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
    
    api_key = input("Enter API key: ").strip()
    
    client = OpenAI(api_key=api_key)
    
    # 1. Ask for the search key
    key = input("Enter search key: ").strip()
    if not key:
        key = ""
    
    # 2. Ask for folder
    print("Please choose a folder...")
    folder = select_folder()
    if not folder:
        print("No folder selected. Exiting.")
        return

    folder_path = Path(folder)

    # 3. Search for matching filenames then append values to ret
    ret = []
    for fname in os.listdir(folder_path):
        if key.lower() in fname.lower():
            full_path = folder_path / fname
            if full_path.is_file():
                path = str(full_path.resolve())
                ocred = ocr_pdf_with_openai(path, client)
                print(ocred)
                ret.append(ocred)


    # 4. Write results to output.csv
    output_path = folder_path / "output.csv"
    
    fieldnames = ret[0].keys()

    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)

        # Write header row
        writer.writeheader()

        # Write each row
        for item in ret:
            writer.writerow(item)

    print(f"\nDone! Saved to: {output_path}")
    input("Press Enter to close")

if __name__ == "__main__":
    main()
    close_terminal_if_frozen()
