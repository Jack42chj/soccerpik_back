from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import json

app = Flask(__name__)
CORS(app, resources={r'/*': {'origins': '*'}}, supports_credentials=True)

@app.route("/", methods=['GET'])
def router():
    print("root url!")
    return jsonify("root url")

@app.route("/search", methods=['POST'])
def search():
    data = request.get_json()
    length = data['length']
    width = data['width']
    height = data['height']
    field = data['field']
    print(data['length'])
    print(data['width'])
    print(data['height'])
    print(data['field'])
    return jsonify({'status': True})

@app.route("/result", methods=['GET'])
def result():
    xlsxname = openpyxl.load_workbook('../footballboots.xlsx', read_only=True)
    sheet = xlsxname.worksheets[0]

    data = {}
    data['result'] = []

    key_list = []
    for i in range(1, sheet.max_column + 1):
        key_list.append(sheet.cell(row=1,column=i).value)

    data_list = {}
    for i in range(2, sheet.max_row + 1):
        temp_dict = {}
        for j in range(1, sheet.max_column + 1):
            val = sheet.cell(row=i, column=j).value
            temp_dict[key_list[j - 1]] = val
        data['result'].append(temp_dict)
    xlsxname.close()
    result = json.dumps(data)
    return result

if __name__ == "__main__":
    app.run(host="0.0.0.0", port="5000")