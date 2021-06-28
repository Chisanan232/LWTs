from spec import FullTestSpec, RegressionTestSpec, EasyTestSpec
from result import ResultReport

from openpyxl import Workbook
import json


# test_spec = FullTestSpec()

# test_spec.loading_content()

# print("++++++++++++++ Full Test +++++++++++++++++")
# test_spec.get_sheet_by_name(name="Full")
# all_modul = test_spec.get_all_device_model()
# print("all_modul: ", all_modul)
# __device_series = test_spec.get_all_test_cases(device_model="COVR-1103_A1")
# print(__device_series)
# __device_series_test_time = test_spec.get_each_cases_testing_time(device_model="COVR-1103_A1")
# print(__device_series_test_time)

# test_spec.close_workbook()

# test_spec = RegressionTestSpec()

# # test_spec.loading_content()

# print("++++++++++++++ Regression Test +++++++++++++++++")
# # test_spec.get_sheet_by_name(name="Regression")
# __device_series = test_spec.get_all_test_cases(device_model="COVR-1103_A1")
# print(__device_series)
# __device_series_test_time = test_spec.get_each_cases_testing_time(device_model="COVR-1103_A1")
# print(__device_series_test_time)

# test_spec.close_workbook()


# test_spec = EasyTestSpec()

# # test_spec.loading_content()

# print("++++++++++++++ Easy Test +++++++++++++++++")
# # test_spec.get_sheet_by_name(name="Easy")
# __device_series = test_spec.get_all_test_cases(device_model="COVR-1103_A1")
# print(__device_series)
# __device_series_test_time = test_spec.get_each_cases_testing_time(device_model="COVR-1103_A1")
# print(__device_series_test_time)

# test_spec.close_workbook()


# [
# {'device_model': '', 
#  'spec_type': 'None', 
#  'items': []}, 
# {'device_model': 'COVR-1103_A1', 
#  'spec_type': 'full_test', 
#  'items': ['{COVR-Reunion: 1}', '{FOTA_In_Wizard: 12.5}', '{Full_Total: =SUM(C2:R2)}']}
# ]


# {"spec_testing_time_0":{"device_model":"COVR-1103_A1","spec_type":"full_test","items":["{COVR-Reunion: 1}","{D-Link_DeFend_full_function: 0}","{FOTA_In_Wizard: 12.5}","{Side_Menu: 0.5}","{mesh_compatibility_: 1.2}","{smart_connect_Wi-Fi_Settings: 3.5}","{wizard_: 4.5}"]},"spec_testing_time_1":{"device_model":"COVR-1103_A1","spec_type":"regression_test","items":["{FOTA_in_wizard: 12.5}","{Remote_management: 7.3}","{mesh_compatibility_: 1.2}"]},"spec_testing_time_2":{"device_model":"COVR-1103_A1","spec_type":"easy_test","items":["{Easy_: 2}"]},"spec_testing_time_3":{"device_model":"DAP-1620","spec_type":"full_test","items":["{Management-AP: 6.7}"]},"spec_testing_time_4":{"device_model":"DAP-1620","spec_type":"regression_test","items":["{Management-AP: 6.7}","{wizard_: 3.5}"]}}

# target_info = {"spec_testing_time_0":{"device_model":"","spec_type":"None","items":[]},"spec_testing_time_1":{"device_model":"COVR-1103_A1","spec_type":"full_test","items":["{COVR-Reunion: 1}","{FOTA_In_Wizard: 12.5}","{Full_Total: =SUM(C2:R2)}"]}}
target_info = {"spec_testing_time_1":{"device_model":"COVR-1103_A1","spec_type":"full_test","items":[{"COVR-Reunion": '1'},{"FOTA_In_Wizard": "12.5"}]}}
workbook: Workbook = ResultReport.general_report(target_info=target_info)
workbook.save("/Users/bryantliu/Downloads/test.xlsx")

