from flask import Flask, request, render_template, send_from_directory

from .export_service import ExportService
from .spec_service import SpecService


def create_app() -> Flask:
    _app = Flask(__name__, template_folder="./static/templates")

    _Spec_Svc = SpecService()
    _Export_Svc = ExportService()


    @_app.route("/", methods=["GET"])
    def welcome() -> str:
        return "Welcom to Leslie Wish Tools - 1"


    @_app.route("/spec-testing-time", methods=["GET"])
    def testing_time_page():
        return render_template("index.html")


    @_app.route("/static/<path:filename>", methods=["GET"])
    def images(filename):
        return send_from_directory('img', filename)
        

    @_app.route("/api/v1/test", methods=["GET"])
    def index():
        return "It's Leslie's Tool!"


    @_app.route("/api/v1/deviceModels", methods=["GET"])
    def device_models():
        _device_modules = _Spec_Svc.get_device_modules()
        return _device_modules


    @_app.route("/api/v1/testing-time", methods=["POST"])
    def each_items_testing_time_by_device_model():
        # # selected_spec = "full_test"
        # # device_model = "COVR-1103_A1"
        _testing_time_info = _Spec_Svc.get_all_testing_info(request=request)
        return _testing_time_info


    @_app.route("/api/v1/calculate-timing", methods=["GET"])
    def calculate_timing():
        items_time = request.form["items_time"]
        sum(items_time)


    @_app.route("/api/v1/export-testing-time-result", methods=["POST"])
    def export_testing_time_result():
        """
        Description:
        I refer tp the answer at stockoverFlow:
        https://stackoverflow.com/questions/42957871/return-a-created-excel-file-with-flask

        :return: A HTTP response which is office excel binary data.
        """
        return _Export_Svc.export(request=request)

    return _app
