from flask import Flask,jsonify,request,send_from_directory
import json
from datetime import datetime
import pandas as pd

app = Flask(__name__)
"""
json data is loaded as global variable to avoid loading for every request
"""
with open("excel_data.json") as fp:
    data = json.loads(fp.read())

@app.route('/total', methods=["GET"])
def get_data():
    """
    day is the input
    given input day is converted into datetime object and compared with dates in the json data
    cummulative sum of total weight, quantity, length is returned

    :return: total data
    """
    day=request.args.get("date")
    query_date = datetime.strptime(day, "%d-%m-%Y")
    total_weight, total_length, total_quantity = float(), float(), float()
    found = False
    for sub_data in data:
        record_date=datetime.strptime(sub_data["DateTime"], "%Y-%m-%dT%H:%M:%SZ")
        if record_date.date()==query_date.date():
            total_length+=sub_data["Length"]
            total_quantity+=sub_data["Quantity"]
            total_weight+=sub_data["Weight"]
            found = True
    if not found:
        return jsonify(data="Not found")
    return jsonify({"totalWeight": total_weight,
                    "totalLength": total_length,
                    "totalQuantity": total_quantity})


@app.route("/excelreport",methods=["GET"])
def get_report():
    """
    segregates the given data into seperate sheets basing on the dates and returns the filtered excel file
    :return: excel file
    """
    df = pd.DataFrame(data)
    dates_list = list()
    print(type(df.DateTime))
    def convert_date(in_date):
        in_date=datetime.strptime(in_date, "%Y-%m-%dT%H:%M:%SZ")
        return in_date.date()

    df.DateTime = df.DateTime.apply(convert_date)
    dates_list = set(df.DateTime)
    with pd.ExcelWriter('output.xlsx') as writer:
        for date_ in dates_list:
          df[df["DateTime"]==date_].to_excel(writer,sheet_name=str(date_))
    return send_from_directory('.',filename="output.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run()
