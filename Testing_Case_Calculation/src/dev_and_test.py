from home_product_spec import FullTestSpec, RegressionTestSpec, EasyTestSpec


test_spec = FullTestSpec()

# test_spec.loading_content()

print("++++++++++++++ Full Test +++++++++++++++++")
# test_spec.get_sheet_by_name(name="Full")
all_modul = test_spec.get_all_device_model()
print("all_modul: ", all_modul)
# __device_series = test_spec.get_all_test_cases(device_model="COVR-1103_A1")
# print(__device_series)
# __device_series_test_time = test_spec.get_each_cases_testing_time(device_model="COVR-1103_A1")
# print(__device_series_test_time)

test_spec.close_workbook()


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
