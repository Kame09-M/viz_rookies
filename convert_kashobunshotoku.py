import openpyxl
import csv

pref_map = {
    '北海道':'北海道', '青\u3000森':'青森県', '岩\u3000手':'岩手県',
    '宮\u3000城':'宮城県', '秋\u3000田':'秋田県', '山\u3000形':'山形県',
    '福\u3000島':'福島県', '茨\u3000城':'茨城県', '栃\u3000木':'栃木県',
    '群\u3000馬':'群馬県', '埼\u3000玉':'埼玉県', '千\u3000葉':'千葉県',
    '東\u3000京':'東京都', '神奈川':'神奈川県', '新\u3000潟':'新潟県',
    '富\u3000山':'富山県', '石\u3000川':'石川県', '福\u3000井':'福井県',
    '山\u3000梨':'山梨県', '長\u3000野':'長野県', '岐\u3000阜':'岐阜県',
    '静\u3000岡':'静岡県', '愛\u3000知':'愛知県', '三\u3000重':'三重県',
    '滋\u3000賀':'滋賀県', '京\u3000都':'京都府', '大\u3000阪':'大阪府',
    '兵\u3000庫':'兵庫県', '奈\u3000良':'奈良県', '和歌山':'和歌山県',
    '鳥\u3000取':'鳥取県', '島\u3000根':'島根県', '岡\u3000山':'岡山県',
    '広\u3000島':'広島県', '山\u3000口':'山口県', '徳\u3000島':'徳島県',
    '香\u3000川':'香川県', '愛\u3000媛':'愛媛県', '高\u3000知':'高知県',
    '福\u3000岡':'福岡県', '佐\u3000賀':'佐賀県', '長\u3000崎':'長崎県',
    '熊\u3000本':'熊本県', '大\u3000分':'大分県', '宮\u3000崎':'宮崎県',
    '鹿児島':'鹿児島県', '沖\u3000縄':'沖縄県'
}

# ① xlsxから所定内給与額を読み込む（月収 = 千円 × 1000）
wb = openpyxl.load_workbook('(1-10-sanko1).xlsx', data_only=True)
ws = wb['男女計']

salary = {}
for row in ws.iter_rows(min_row=11, values_only=True):
    raw_name = row[2]  # index=2: 都道府県名
    wage = row[9]      # index=9: 所定内給与額（千円）
    if raw_name is None or wage is None:
        continue
    pref = pref_map.get(str(raw_name).strip())
    if pref:
        salary[pref] = int(float(wage) * 1000)

# ② CSVから2022年度の消費支出を読み込む
consumption = {}
with open('Viz Rookies利用データ_SSDSE-県別推移.csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row['年度'] != '2022':
            continue
        pref = row['都道府県']
        val = row['消費支出（二人以上の世帯）']
        if val:
            consumption[pref] = int(float(val))

# ③ 結合して可処分所得を計算
out_rows = []
for pref in sorted(salary.keys()):
    monthly = salary[pref]
    cons = consumption.get(pref)
    if cons is None:
        print(f'WARNING: 消費支出データなし -> {pref}')
        continue
    disposable = monthly - cons
    out_rows.append([pref, monthly, cons, disposable])

with open('kashobunshotoku.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['都道府県', '月収（円）', '消費支出（円/月）', '可処分所得（円/月）'])
    writer.writerows(out_rows)

print(f'完了: {len(out_rows)} 行を出力しました')

# 確認表示
print('\n--- 確認 ---')
for row in out_rows:
    if row[0] in ('東京都', '宮城県', '福島県'):
        print(row)
