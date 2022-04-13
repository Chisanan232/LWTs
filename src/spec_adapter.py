from flask import jsonify
from openpyxl.workbook import Workbook

from .spec import TestSpec



class DLinkHomeProductSpec:

    def __init__(self, test_spec: TestSpec):
        self.test_spec = test_spec


    @property
    def current_sheet_page(self) -> Workbook:
        return self.test_spec.current_sheet_page


    def get_each_items_testing_time(self, device_model: str):
        each_case_testing_time = self.test_spec.get_each_cases_testing_time(device_model=device_model)
        self.test_spec.close_workbook()
        return jsonify(each_case_testing_time)

