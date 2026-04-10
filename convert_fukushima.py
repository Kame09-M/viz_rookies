import openpyxl
import csv

# 47都道府県コード（01000〜47000）
PREF_CODES = {f'{i:02d}000' for i in range(1, 48)}

# 抽出する年齢階級と列インデックス（0-indexed）
# 総数: idx 9=総数, 10=0-4, 11=5-9, 12=10-14, 13=15-19, 14=20-24, 15=25-29, 16=30-34
# 男:   idx 29=総数, 30=0-4, 31=5-9, 32=10-14, 33=15-19, 34=20-24, 35=25-29, 36=30-34
# 女:   idx 49=総数, 50=0-4, 51=5-9, 52=10-14, 53=15-19, 54=20-24, 55=25-29, 56=30-34
AGE_COLS = {
    '総数': {'15-19歳': 13, '20-24歳': 14, '25-29歳': 15, '30-34歳': 16},
    '男':   {'15-19歳': 33, '20-24歳': 34, '25-29歳': 35, '30-34歳': 36},
    '女':   {'15-19歳': 53, '20-24歳': 54, '25-29歳': 55, '30-34歳': 56},
}

wb = openpyxl.load_workbook('b00802s.xlsx', data_only=True)
ws = wb.active

rows_out = []
for row in ws.iter_rows(min_row=8, values_only=True):
    src_code  = row[1]   # 移動前の住所地コード
    nat_code  = row[3]   # 国籍コード
    dest_code = str(row[7]) if row[7] is not None else ''
    dest_name = row[8]   # 移動後の住所地

    # フィルタ：福島県 / 移動者 / 47都道府県
    if src_code != '07000':
        continue
    if nat_code != 60000:
        continue
    if dest_code not in PREF_CODES:
        continue

    for gender, age_map in AGE_COLS.items():
        for age_label, col_idx in age_map.items():
            val = row[col_idx]
            rows_out.append([dest_name, gender, age_label, val if val is not None else 0])

with open('fukushima_destination.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['転出先都道府県', '性別', '年齢階級', '転出者数'])
    writer.writerows(rows_out)

print(f'完了: {len(rows_out)} 行を出力しました')
