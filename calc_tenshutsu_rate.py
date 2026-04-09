import openpyxl
import csv

# --- 1. idou_tableau_v2.csv から女性20〜24歳の転入超過数を取得 ---
nyuucho = {}  # 都道府県 -> 転入超過数（int）

with open('idou_tableau_v2.csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if (row['種別'] == '転入超過数'
                and row['性別'] == '女'
                and row['年齢階級'] == '20歳～24歳'):
            pref = row['都道府県']
            # 全国・3大都市圏・21大都市を除外し、47都道府県のみ残す
            nyuucho[pref] = int(row['人数'])

# --- 2. b02_03.xlsx から国籍総数・女・都道府県の20〜24歳人口を取得 ---
jinko = {}  # 都道府県 -> 20〜24歳女性人口（int）

wb = openpyxl.load_workbook('b02_03.xlsx', read_only=True, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))

for row in rows[11:]:  # ヘッダー行をスキップ
    kokuseki = row[0]
    seibetsu = row[1]
    chiiki_code = row[2]
    chiiki_name = row[3]  # 例: '01000_北海道'

    if kokuseki != '0_国籍総数':
        continue
    if seibetsu != '2_女':
        continue
    if chiiki_code != 'a':  # 'a' = 全国・都道府県のみ
        continue
    if chiiki_name and chiiki_name.startswith('00000_'):  # 全国を除外
        continue

    if chiiki_name:
        # '01000_北海道' → '北海道'
        pref_name = chiiki_name.split('_', 1)[1]
        pop_20_24 = row[9]  # 05_20～24歳 列（0-indexed: 4=総数, 5=0~4歳, ..., 9=20~24歳）
        if pop_20_24 is not None:
            jinko[pref_name] = int(pop_20_24)

wb.close()

# --- 3. 47都道府県で突合し転出超過率を計算 ---
# idou側にある都道府県から全国・3大都市圏・21大都市を除き47都道府県に絞る
PREFS_47 = set(jinko.keys())

results = []
unmatched = []

for pref, pop in sorted(jinko.items()):
    if pref not in nyuucho:
        unmatched.append(pref)
        continue
    nyuucho_val = -nyuucho[pref]                        # 転出超過数 = -(転入超過数)
    rate_pct = nyuucho_val / pop * 100 if pop else None  # 転出超過率(%)
    results.append({
        '都道府県': pref,
        '転出超過数': nyuucho_val,
        '女性20〜24歳人口': pop,
        '転出超過率(%)': round(rate_pct, 4) if rate_pct is not None else '',
    })

if unmatched:
    print(f'警告: 突合できなかった都道府県: {unmatched}')

# --- 4. CSV出力 ---
output_file = 'tenshutsu_rate.csv'
with open(output_file, 'w', newline='', encoding='utf-8') as f:
    fieldnames = ['都道府県', '転出超過数', '女性20〜24歳人口', '転出超過率(%)']
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(results)

print(f'出力完了: {output_file}')
print(f'都道府県数: {len(results)}件')
print()
print('先頭5件:')
for r in results[:5]:
    print(f"  {r['都道府県']}: 転出超過数={r['転出超過数']:,}, 人口={r['女性20〜24歳人口']:,}, 率={r['転出超過率(%)']}%")
