from flask import Flask, request, send_file, jsonify
import openpyxl
from openpyxl.styles import PatternFill
import urllib.request
import json
import io

app = Flask(__name__)

RED_FILL = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

@app.route('/rate')
def get_rate():
    url = 'https://api.exchangerate-api.com/v4/latest/USD'
    with urllib.request.urlopen(url) as response:
        data = json.loads(response.read())
    rate = data['rates']['TJS']
    return jsonify({'rate': rate})

@app.route('/colorize', methods=['POST'])
def colorize():
    file = request.files['file']
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    header_row = [cell.value for cell in ws[1]]
    color_col = None
    for i, h in enumerate(header_row):
        if h == '_rowColor':
            color_col = i + 1
            break

    if color_col:
        for row in ws.iter_rows(min_row=2):
            color_val = row[color_col - 1].value
            if color_val == 'red':
                for cell in row:
                    if cell.column != color_col:
                        cell.fill = RED_FILL
            elif color_val == 'yellow':
                for cell in row:
                    if cell.column != color_col:
                        cell.fill = YELLOW_FILL
        ws.delete_cols(color_col)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='reestr_checked.xlsx'
    )

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
