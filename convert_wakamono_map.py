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

TARGET_AGES = {'15歳～19歳', '20歳～24歳', '25歳～29歳', '30歳～34歳'}

totals = {}

with open('idou_tableau_v2.csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row['種別'] != '転入超過数':
            continue
        if row['性別'] != '総数':
            continue
        if row['年齢階級'] not in TARGET_AGES:
            continue
        if row['都道府県'] not in PREFS:
            continue
        pref = row['都道府県']
        totals[pref] = totals.get(pref, 0) + int(row['人数'])

out_rows = sorted(totals.items(), key=lambda x: x[1])

with open('wakamono_map_15to34.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['都道府県', '転入超過数_15〜34歳'])
    writer.writerows(out_rows)

print(f'総行数: {len(out_rows)}')
print('\nワースト5位（転出超過が最も大きい順）:')
for i, (pref, val) in enumerate(out_rows[:5], 1):
    print(f'  {i}位: {pref} {val:,}')

fukushima_rank = next(i+1 for i, (p, _) in enumerate(out_rows) if p == '福島県')
fukushima_val = totals['福島県']
print(f'\n福島県: {fukushima_rank}位 ({fukushima_val:,})')
