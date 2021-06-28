from spec_adapter import DLinkHomeProductSpec
from spec import TestSpec, ResultTestSpec, FullTestSpec, RegressionTestSpec, EasyTestSpec, DeviceModelDoesnotExistException
from result import ResultReport

from flask import Flask, request, Response, jsonify, render_template, send_from_directory
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook
from datetime import datetime
import re

app = Flask(__name__)
app.config["debug"] = True



@app.route("/spec-testing-time")
def testing_time_page():
    return render_template("index.html")



@app.route("/static/<path:filename>", methods=["GET"])
def images(filename):
    return send_from_directory('img', filename)
    


@app.route("/api/v1/test")
def index():
    return "It's Leslie's Tool!"



@app.route("/api/v1/deviceModels", methods=["GET"])
def device_models():
    target_spec = ResultTestSpec()
    # target_spec = FullTestSpec()
    test = target_spec.get_all_device_model()
    print(f"test value: {test}")
    return jsonify(target_spec.get_all_device_model())



@app.route("/api/v1/testing-time", methods=["POST"])
def each_items_testing_time_by_device_model():
    # selected_spec = "full_test"
    # device_model = "COVR-1103_A1"
    print("request: ", request.form)
    if "spec" in request.form and "deviceModel" in request.form:
        selected_spec = request.form["spec"]
        device_model = request.form["deviceModel"]
    else:
        return jsonify({
            "state": "error",
            "message": "Please use this API with 2 needed parameters 'spec' and 'deviceModel'."
        })
    
    full_test = re.search(re.escape(selected_spec), "full_test", re.IGNORECASE)
    regression_test = re.search(re.escape(selected_spec), "regression_test", re.IGNORECASE)
    easy_test = re.search(re.escape(selected_spec), "easy_test", re.IGNORECASE)
    if full_test is not None:
        target_spec = FullTestSpec()
    elif regression_test is not None:
        target_spec = RegressionTestSpec()
    elif easy_test is not None:
        target_spec = EasyTestSpec()
    else:
        return jsonify({
            "state": "error",
            "message": "Parameter 'spec' is incorrect. Only support 3 value 'full_test', 'regression_test' or 'easy_test'."
        })

    router_spec = DLinkHomeProductSpec(test_spec=target_spec)
    try:
        return router_spec.get_each_items_testing_time(device_model=device_model)
    except DeviceModelDoesnotExistException as e:
        return jsonify({
            "state": "error",
            "message": str(e)
        })



@app.route("/api/v1/export-testing-time-result", methods=["POST"])
def export_testing_time_result():
    """
    Description:
    I refer tp the answer at stockoverFlow:
    https://stackoverflow.com/questions/42957871/return-a-created-excel-file-with-flask

    :return: A HTTP response which is office excel binary data.
    """
    target_info = request.form["target"]
    workbook = ResultReport.general_report(target_info=target_info)
    general_report_datetime = datetime.now().isoformat(timespec='seconds').split("T")[0]
    return Response(
        save_virtual_workbook(workbook=workbook),
        headers={
            'Content-Disposition': f'attachment; filename=D-Link_Wi-Fi_Testing_Time_Report_{general_report_datetime}.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )


if __name__ == '__main__':

    app.run(host="0.0.0.0", port=5223, debug=True)
