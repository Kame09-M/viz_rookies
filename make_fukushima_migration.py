import csv
import io

# エンコーディング自動判定（候補順に試す）
encodings = ["cp932", "shift-jis", "utf-8-sig", "utf-8"]
encoding_used = None
raw_lines = None

for enc in encodings:
    try:
        with open("転出2010-2025.csv", encoding=enc, errors="strict") as f:
            raw_lines = f.read()
        encoding_used = enc
        break
    except Exception:
        continue

if encoding_used is None:
    raise RuntimeError("エンコーディングを判定できませんでした")

print(f"読み込みエンコーディング: {encoding_used}")

# CSV パース（ヘッダー行を探す）
reader = csv.reader(io.StringIO(raw_lines))
rows = list(reader)

# データ行の開始位置を探す（国籍 コード が先頭の行）
header_idx = None
for i, row in enumerate(rows):
    if row and row[0].strip() == "国籍 コード":
        header_idx = i
        break

if header_idx is None:
    raise RuntimeError("ヘッダー行が見つかりません")

print(f"ヘッダー行: {header_idx + 1}行目")

data_rows = rows[header_idx + 1:]

# 抽出条件
# 列インデックス（0始まり）
#   0: 国籍 コード
#   5: 時間軸（年次）  例: "2015年"
#   6: 地域 コード
#  16: 他都道府県への転出者数 総数

results = []
for row in data_rows:
    if len(row) < 17:
        continue
    kokuseki_code = row[0].strip()
    year_str = row[5].strip()          # "2015年"
    chiiki_code = row[6].strip()
    tenshtsu_total = row[16].strip()   # "31,552" など

    # 国籍コード = 60000（全体）
    if kokuseki_code != "60000":
        continue
    # 地域コード = 07000（福島県）
    if chiiki_code != "07000":
        continue
    # 年次 2015〜2024
    if not year_str.endswith("年"):
        continue
    try:
        year = int(year_str.replace("年", ""))
    except ValueError:
        continue
    if not (2015 <= year <= 2024):
        continue

    # カンマ除去して数値化
    try:
        ninzu = int(tenshtsu_total.replace(",", ""))
    except ValueError:
        continue

    results.append({"年": year, "指標": "転出数", "人数": ninzu})

# 移住者数データ（福島県公式）を手動追加
ijusha_data = [
    {"年": 2020, "指標": "移住者数", "人数": 1116},
    {"年": 2021, "指標": "移住者数", "人数": 2333},
    {"年": 2022, "指標": "移住者数", "人数": 2832},
    {"年": 2023, "指標": "移住者数", "人数": 3419},
    {"年": 2024, "指標": "移住者数", "人数": 3799},
]
results.extend(ijusha_data)

# 年→指標 の順でソート（転出数先、移住者数後、年昇順）
order = {"転出数": 0, "移住者数": 1}
results.sort(key=lambda x: (order.get(x["指標"], 99), x["年"]))

print(f"\n抽出件数: {len(results)} 行")
print("年,指標,人数")
for r in results:
    print(f"{r['年']},{r['指標']},{r['人数']}")

# UTF-8 BOM付きで出力
output_path = "fukushima_migration.csv"
with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=["年", "指標", "人数"])
    writer.writeheader()
    writer.writerows(results)

print(f"\n保存完了: {output_path}")
