import openpyxl
import csv

# --- 年齢階級マッピング: idou → b02_03.xlsx の列インデックス(0-based) ---
# b02_03: col4=総数, col5=0~4歳, ..., col23=90~94歳, col24=95~99歳, ...col27=110歳以上
# 90歳以上はcol23~col27の合計
AGE_COL_MAP = {
    '0歳～4歳':   [5],
    '5歳～9歳':   [6],
    '10歳～14歳': [7],
    '15歳～19歳': [8],
    '20歳～24歳': [9],
    '25歳～29歳': [10],
    '30歳～34歳': [11],
    '35歳～39歳': [12],
    '40歳～44歳': [13],
    '45歳～49歳': [14],
    '50歳～54歳': [15],
    '55歳～59歳': [16],
    '60歳～64歳': [17],
    '65歳～69歳': [18],
    '70歳～74歳': [19],
    '75歳～79歳': [20],
    '80歳～84歳': [21],
    '85歳～89歳': [22],
    '90歳以上':   [23, 24, 25, 26, 27],  # 90~94, 95~99, 100~104, 105~109, 110歳以上
}

# 性別マッピング: idou → b02_03
GENDER_MAP = {
    '総数': '0_総数',
    '男':   '1_男',
    '女':   '2_女',
}

# --- 1. b02_03.xlsx から47都道府県の人口を取得 ---
# key: (都道府県名, b02性別, 年齢列インデックス) → 人口
# まず都道府県ごとの行データを辞書で保持
# structure: jinko[都道府県名][b02性別] = row（全列）
jinko = {}  # {pref: {gender_b02: row}}
pref_47 = []  # 47都道府県名リスト（順序保持）

wb = openpyxl.load_workbook('b02_03.xlsx', read_only=True, data_only=True)
ws = wb.active
rows_xlsx = list(ws.iter_rows(values_only=True))

for row in rows_xlsx[11:]:
    kokuseki   = row[0]
    seibetsu   = row[1]
    chiiki_cd  = row[2]
    chiiki_nm  = row[3]

    if kokuseki != '0_国籍総数':
        continue
    if chiiki_cd != 'a':        # 都道府県・全国のみ
        continue
    if not chiiki_nm or chiiki_nm.startswith('00000_'):  # 全国を除外
        continue

    pref_name = chiiki_nm.split('_', 1)[1]  # '01000_北海道' → '北海道'

    if pref_name not in jinko:
        jinko[pref_name] = {}
        pref_47.append(pref_name)

    jinko[pref_name][seibetsu] = row  # seibetsu: '0_総数','1_男','2_女'

wb.close()
print(f'b02_03.xlsx: {len(pref_47)}都道府県を読み込み')

# --- 2. idou_tableau_v2.csv から転入超過数を取得 ---
# key: (都道府県, 性別, 年齢階級) → 転入超過数
idou = {}

with open('idou_tableau_v2.csv', encoding='utf-8') as f:
    for row in csv.DictReader(f):
        if row['種別'] != '転入超過数':
            continue
        if row['年齢階級'] == '総数':
            continue
        pref   = row['都道府県']
        gender = row['性別']      # 総数/男/女
        age    = row['年齢階級']  # 0歳～4歳 など
        val    = int(row['人数'])
        idou[(pref, gender, age)] = val

print(f'idou_tableau_v2.csv: 読み込み完了')

# --- 3. 突合して転出超過率を計算 ---
results = []

for pref in pref_47:
    for idou_gender, b02_gender in GENDER_MAP.items():
        if b02_gender not in jinko[pref]:
            continue
        pop_row = jinko[pref][b02_gender]

        for age_label, col_idxs in AGE_COL_MAP.items():
            key = (pref, idou_gender, age_label)
            if key not in idou:
                continue

            # 転出超過数 = -(転入超過数)
            tenshutsu_cho = -idou[key]

            # 人口: 対応列の合計（90歳以上は複数列を合計）
            def to_int(v):
                try: return int(v)
                except (TypeError, ValueError): return 0
            pop = sum(to_int(pop_row[i]) for i in col_idxs)

            rate = round(tenshutsu_cho / pop * 100, 4) if pop else ''

            results.append({
                '都道府県':       pref,
                '性別':           idou_gender,
                '年齢階級':       age_label,
                '転出超過数':     tenshutsu_cho,
                '人口':           pop,
                '転出超過率(%)':  rate,
            })

print(f'レコード数: {len(results):,}件')

# --- 4. CSV出力 ---
output_file = 'tenshutsu_all.csv'
with open(output_file, 'w', newline='', encoding='utf-8') as f:
    fieldnames = ['都道府県', '性別', '年齢階級', '転出超過数', '人口', '転出超過率(%)']
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(results)

print(f'出力完了: {output_file}')
print()
print('先頭5件:')
for r in results[:5]:
    print(f"  {r['都道府県']} {r['性別']} {r['年齢階級']}: "
          f"転出超過数={r['転出超過数']:,}, 人口={r['人口']:,}, 率={r['転出超過率(%)']}%")
