import csv

PREFS = {
    '北海道','青森県','岩手県','宮城県','秋田県','山形県','福島県',
    '茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県',
    '新潟県','富山県','石川県','福井県','山梨県','長野県','岐阜県',
    '静岡県','愛知県','三重県','滋賀県','京都府','大阪府','兵庫県',
    '奈良県','和歌山県','鳥取県','島根県','岡山県','広島県','山口県',
    '徳島県','香川県','愛媛県','高知県','福岡県','佐賀県','長崎県',
    '熊本県','大分県','宮崎県','鹿児島県','沖縄県'
}

def load_data(filepath, target_indicator):
    """都道府県→{年次: 値} の辞書を返す"""
    data = {}
    with open(filepath, encoding='shift-jis') as f:
        reader = csv.reader(f)
        next(reader)  # skiprows=1
        for row in reader:
            if row[1] != target_indicator:  # 出生数・婚姻件数のみ（率は除外）
                continue
            pref = row[3]
            year = row[5]
            val  = row[7]
            if pref not in PREFS:
                continue
            if year not in ('2018年', '2023年'):
                continue
            if pref not in data:
                data[pref] = {}
            try:
                data[pref][year] = int(float(val))
            except ValueError:
                pass  # '…' など欠損値はスキップ
    return data

birth    = load_data('FEH_00450011_260413140732.csv', '出生数')
marriage = load_data('FEH_00450011_260413141613.csv', '婚姻件数')

# 結合・計算
rows = []
for pref in PREFS:
    b18 = birth[pref]['2018年']
    b23 = birth[pref]['2023年']
    b_rate = round((b23 - b18) / b18 * 100, 1)

    m18 = marriage[pref]['2018年']
    m23 = marriage[pref]['2023年']
    m_rate = round((m23 - m18) / m18 * 100, 1)

    rows.append({'都道府県': pref,
                 '出生数_2018': b18, '出生数_2023': b23, '出生減少率(%)': b_rate,
                 '婚姻件数_2018': m18, '婚姻件数_2023': m23, '婚姻減少率(%)': m_rate})

# 順位付け（減少率が小さい＝最も減少した県が1位）
rows.sort(key=lambda x: x['出生減少率(%)'])
for i, r in enumerate(rows, 1):
    r['出生減少率順位'] = i

rows.sort(key=lambda x: x['婚姻減少率(%)'])
for i, r in enumerate(rows, 1):
    r['婚姻減少率順位'] = i

# 都道府県名順に並べ直して出力
rows.sort(key=lambda x: x['出生減少率順位'])

with open('birth_marriage_ranking.csv', 'w', newline='', encoding='utf-8') as f:
    fieldnames = ['都道府県','出生数_2018','出生数_2023','出生減少率(%)',
                  '出生減少率順位','婚姻件数_2018','婚姻件数_2023',
                  '婚姻減少率(%)','婚姻減少率順位']
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)

print(f'完了: {len(rows)} 行を出力しました')

# 確認表示
birth_sorted   = sorted(rows, key=lambda x: x['出生減少率順位'])
marriage_sorted = sorted(rows, key=lambda x: x['婚姻減少率順位'])

print('\n出生減少率ワースト5:')
for r in birth_sorted[:5]:
    print(f'  {r["出生減少率順位"]}位: {r["都道府県"]}  {r["出生減少率(%)"]:.1f}%')

print('\n婚姻減少率ワースト5:')
for r in marriage_sorted[:5]:
    print(f'  {r["婚姻減少率順位"]}位: {r["都道府県"]}  {r["婚姻減少率(%)"]:.1f}%')

fuku = next(r for r in rows if r['都道府県'] == '福島県')
print(f'\n福島県:')
print(f'  出生  : {fuku["出生減少率順位"]}位  {fuku["出生減少率(%)"]:.1f}%  ({fuku["出生数_2018"]:,}人 → {fuku["出生数_2023"]:,}人)')
print(f'  婚姻  : {fuku["婚姻減少率順位"]}位  {fuku["婚姻減少率(%)"]:.1f}%  ({fuku["婚姻件数_2018"]:,}件 → {fuku["婚姻件数_2023"]:,}件)')
