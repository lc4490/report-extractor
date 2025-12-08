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

SYSTEM_PROMPT = """你是一個精準的資料擷取引擎，專門處理紡織／布料的「內部檢驗報告」。

你的任務是從圖像中擷取所有檢驗結果，不是只有第一筆。

⚠️ **最重要規則：**
你輸出的內容必須是 **一個完整且有效的 JSON 物件**（不能有任何多餘文字、不能用 Markdown、不能有說明）。

---

# 🚩 **輸出格式（務必遵守）**

你必須輸出以下結構的 JSON（所有值皆為字串；多筆值使用 array）：

```json
{
  "訂單編號": "",
  "重量": "",
  "厚度": "",
  "拉力強度_warp": [],
  "拉力強度_weft": [],
  "剝離強度_warp": [],
  "剝離強度_weft": [],
  "撕裂強度_warp": [],
  "撕裂強度_weft": [],
  "高週波強度B/B_warp": [],
  "高週波強度B/B_weft": [],
  "高迪波強度F/B_warp": [],
  "高迪波強度F/B_weft": []
}
```

### 每個欄位規則：

* 所有 measurement 欄位都是 **list of strings**
* 每一筆代表「檢驗結果」表格中的一行（例如 支數 1、2、3）
* 若該項目完全不存在，該欄位輸出 `[]`
* 若某一格顯示 `ND`，請輸出 `"ND"`
* 若該格完全空白，輸出 `"N/A"`

---

# 🚩 **資料擷取規則**

## 1. 表頭欄位

從報告最上方擷取：

* `"訂單編號"`（如 24072201-3、S25092202-2）
* `"重量"`（只保留數字，例如 221.5）
* `"厚度"`（例如 0.31）

不要保留單位（g/m2、mm）。

---

## 2. measurement 欄位（最關鍵）

你必須找到對應的表格，並擷取 **所有「檢驗結果」的行**。

永遠 **忽略 標準／規格／試驗標準 列**。

多筆資料例：

```
支數 | warp | weft
1    | 215.5 | 187.4
2    | 274.5 | 180.5
3    | 244.0 | 172.2
```

→ 你必須輸出：

```
"高週波強度B/B_warp": ["215.5","274.5","244.0"],
"高週波強度B/B_weft": ["187.4","180.5","172.2"],
```

---

## 3. 各表格對應方式

### (a) 拉力強度 (N/in)

標題包含：`拉力強度`
→ 擷取所有檢驗結果行（warp & weft）

### (b) 剝離強度 (N/in)

標題包含：`剝離強度`
→ 多筆列全部輸出

### (c) 撕裂強度 (N)

標題包含：`撕裂強度`
→ 多筆列全部輸出

### (d) 高週波強度 B/B (N/in)

標題包含：`B/B` 或 `高週波強度(N/in)-B/B`
→ 多筆列全部輸出
→ 不能取標準值的 220.0 / 180.0 行

### (e) 高迪波強度 F/B (N/in)

標題包含：`F/B`
→ 多筆列全部輸出
→ 若整份報告沒有 F/B 表格 → 輸出空 array (`[]`)

---

# 🚩 **數值清理**

1. 移除符號（例如 `215.5*` → `"215.5"`）
2. 保留小數格式原樣
3. 若表格單元格顯示 `ND` → `"ND"`
4. 若單元格完全是空白 → `"N/A"`

---

# 🚩 **最後規則（務必遵守）**

* 你只能輸出 **一個 JSON 物件**
* 首字必須是 `{`
* 末字必須是 `}`
* 中間所有 key 的內容不可缺漏
* 所有非必填項目若找不到 → 使用空 array `[]`
* 不可輸出 Markdown 或任何說明文字
        """

array_fields = [
        "拉力強度_warp", "拉力強度_weft",
        "剝離強度_warp", "剝離強度_weft",
        "撕裂強度_warp", "撕裂強度_weft",
        "高週波強度B/B_warp", "高週波強度B/B_weft",
        "高迪波強度F/B_warp", "高迪波強度F/B_weft",
    ]
# select folder
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select a folder to search")
    root.destroy()
    return folder

# takes folder path, key, and openai api key and parses all files that match key. then, it sends it these pdf files to get ocred, and returns the results
def collect_results(folder_path: Path, key: str, client: OpenAI) -> list[dict]:
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
        file_rows = build_rows_for_file(full_path, client)
        if file_rows:
            results.extend(file_rows)

    return results


# takes a file path and an openai api key and sends it to get ocred, then cleans up the results to get appended
def build_rows_for_file(path: Path, client: OpenAI) -> list[dict]:
    data = ocr_pdf_with_openai(str(path.resolve()), client)
    num_rows = max(len(data[field]) for field in array_fields)

    rows = []
    for i in range(num_rows):
        row = {
            "訂單編號": data["訂單編號"],
            "重量": data["重量"],
            "厚度": data["厚度"],
            "roll": i + 1,
        }
        for field in array_fields:
            vals = data[field]
            if not vals:
                row[field] = "N/A"
            elif i < len(vals):
                row[field] = vals[i]
            else:
                row[field] = vals[-1]
        rows.append(row)
    return rows

# takes a pdf file path, and returns some raw ocr 
def ocr_pdf_with_openai(pdf_path: Path, client):
    """
    Convert first page of PDF to image, send to OpenAI Vision,
    and return ONLY numbers (including -, /, .).
    """
    print(pdf_path)
    pages = convert_from_path(str(pdf_path), dpi=300)

    # Take first page for now
    page = pages[0]
    img_b64 = pil_to_b64(page)

    system_prompt = SYSTEM_PROMPT

    response = client.chat.completions.create(
        model="gpt-4o",
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
        data = parse_model_json(content)
    except Exception as e:
        print("Model did not return valid JSON. Raw content below:\n", e)
        print(content)
        raise  # don't continue with invalid data
    
    return data

# helper function for ocr
def parse_model_json(raw: str) -> dict:
    if "```" in raw:
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1:
            raise ValueError("Could not find JSON object braces in model output.")
        raw = raw[start:end+1]

    data = json.loads(raw)

    # If the model ever wraps the object in a list, unwrap [ {...} ]
    if isinstance(data, list):
        if len(data) == 1 and isinstance(data[0], dict):
            data = data[0]
        else:
            raise ValueError(f"Expected a single JSON object, got list: {data}")

    if not isinstance(data, dict):
        raise ValueError(f"Expected JSON object, got {type(data)}")

    return data

# helper function for ocr
def pil_to_b64(img):
    """Convert PIL image to base64 PNG."""
    buf = BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")

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



def main():
    print("=== Report Extractor ===")
    
    # 1. Ask user for API key
    api_key = input("Enter API key: ").strip()
    client = OpenAI(api_key=api_key)
    
    # 2. Ask for the search key
    key = input("Enter search key: ").strip()
    if not key:
        key = ""
    
    # 3. Ask for folder
    print("Please choose a folder...")
    folder = select_folder()
    if not folder:
        print("No folder selected. Exiting.")
        return

    folder_path = Path(folder)

    # 4. Search for matching filenames, run OCR, then parse through them, append them to ret
    ret = collect_results(folder_path, key, client)

    # 5. Write results to output.csv
    output_path = folder_path / "output.csv"
    
    if ret:
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
