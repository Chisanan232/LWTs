from flask.wrappers import Request
from flask import jsonify
import re

from .spec_adapter import DLinkHomeProductSpec
from .spec import ResultTestSpec, FullTestSpec, RegressionTestSpec, EasyTestSpec, DeviceModelDoesnotExistException



class SpecService:

    def __init__(self) -> None:
        self._target_spec = ResultTestSpec()


    def get_device_modules(self):
        test = self._target_spec.get_all_device_model()
        print(f"test value: {test}")
        return jsonify(self._target_spec.get_all_device_model())
    

    def get_all_testing_info(self, request: Request):
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

