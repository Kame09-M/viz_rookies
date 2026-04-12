import openpyxl
import csv
import re

wb = openpyxl.load_workbook('b02_03.xlsx', data_only=True)
ws = wb['b02_03']

# 都道府県コード: 01000〜47000
PREF_CODE_RE = re.compile(r'^(0[1-9]|[1-3][0-9]|4[0-7])000_')

out_rows = []

for row in ws.iter_rows(min_row=12, values_only=True):
    if row[0] != '0_国籍総数':
        continue
    if row[1] != '0_総数':
        continue
    if row[2] != 'a':
        continue
    region = str(row[3]) if row[3] else ''
    if not PREF_CODE_RE.match(region):
        continue

    pref_name = region.split('_', 1)[1]  # 例: 07000_福島県 → 福島県
    total_pop  = row[4]   # 00_総数
    age_15_19  = row[8]   # 04_15〜19歳
    age_20_24  = row[9]   # 05_20〜24歳
    age_25_29  = row[10]  # 06_25〜29歳
    age_30_34  = row[11]  # 07_30〜34歳

    young_pop = age_15_19 + age_20_24 + age_25_29 + age_30_34
    ratio = round(young_pop / total_pop * 100, 2)

    out_rows.append([pref_name, total_pop, young_pop, ratio])

with open('wakamono_population_15to34.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['都道府県', '総人口', '15〜34歳人口', '15〜34歳比率(%)'])
    writer.writerows(out_rows)

print(f'完了: {len(out_rows)} 行を出力しました')

# 福島県の行と順位（比率の降順）
ranked = sorted(out_rows, key=lambda x: x[3], reverse=True)
for rank, row in enumerate(ranked, 1):
    if row[0] == '福島県':
        print(f'\n福島県: {rank}位 / 47都道府県')
        print(f'  総人口         : {row[1]:,}人')
        print(f'  15〜34歳人口   : {row[2]:,}人')
        print(f'  15〜34歳比率   : {row[3]}%')
        break
