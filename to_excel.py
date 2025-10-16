from flask import Flask, request
from openpyxl import Workbook, load_workbook
import os
from flask_cors import CORS  # 追加
from datetime import datetime  # 追加

app = Flask(__name__)
CORS(app)  # 追加

EXCEL_PATH = 'test.xlsx'

@app.route('/submit', methods=['POST'])
def submit():
    # 変更: HTMLフォームのname属性に合わせてデータ取得
    data = [
        request.form.get('address'),            # name="address"
        request.form.get('org_name'),           # name="org_name"
        request.form.get('person_name'),        # name="person_name"
        request.form.get('tel'),                # name="tel"
        request.form.get('fax'),                # name="fax"
        request.form.get('dispatch_time'),      # name="dispatch_time"
        request.form.get('dest_address'),       # name="dest_address"
        request.form.get('dest_place'),         # name="dest_place"
        request.form.get('dest_tel'),           # name="dest_tel"
        request.form.get('dest_fax'),           # name="dest_fax"
        request.form.get('reason')              # name="reason"
    ]

    # 変更: test.xlsxに書き込み（A列:行番号, B列:日時, C列以降:データ）
    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        # 変更: ヘッダー行をA列・B列・C列以降で作成
        ws.append(
            ['行番号', '書き込み日時', '住所', '名前（団体名）', '担当者名', '電話番号', 'FAX', '派遣日時', '派遣先の住所', '派遣先の会場名', '派遣先の電話番号', '派遣先のFAX', '申請事由']
        )
    else:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active

    # 変更: 行番号（A列）と日時（B列）を追加
    row_number = ws.max_row  # ヘッダー行を含むので、次の行番号
    write_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ws.append([row_number, write_time] + data)
    wb.save(EXCEL_PATH)
    return '', 200

if __name__ == '__main__':
    # 変更: localhost:5000で起動
    app.run(host='localhost', port=5000)