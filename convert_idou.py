import openpyxl
import csv

files = [
    ('a00301s.xlsx', '転入者数'),
    ('a00302s.xlsx', '転出者数'),
    ('a00303s.xlsx', '転入超過数'),
]

all_rows = []

for fname, category in files:
    wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Row 4 (index 3): 性別ヘッダー, Row 5 (index 4): 年齢ヘッダー
    gender_row = rows[3]
    age_row = rows[4]

    # 列7〜66（0-indexed: 7〜66）に性別×年齢の組み合わせが60列
    col_headers = []
    for i in range(7, 67):
        col_headers.append((gender_row[i], age_row[i]))

    # データは行7（index 6）以降、国籍コード60000（移動者）のみ
    for row in rows[6:]:
        if row[6] is None:
            continue
        if row[1] != 60000:  # 国籍コード: 60000=移動者のみ残す
            continue
        prefecture = row[6]
        year = row[4]

        for j, (gender, age) in enumerate(col_headers):
            value = row[7 + j]
            all_rows.append({
                '種別': category,
                '都道府県': prefecture,
                '年次': year,
                '性別': gender,
                '年齢階級': age,
                '人数': value,
            })

    wb.close()
    print(f'{fname}: 処理完了')

output_file = 'idou_tableau_v2.csv'
with open(output_file, 'w', newline='', encoding='utf-8') as f:
    fieldnames = ['種別', '都道府県', '年次', '性別', '年齢階級', '人数']
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(all_rows)

print(f'\n出力完了: {output_file}')
print(f'総行数: {len(all_rows):,} 行（ヘッダー除く）')
