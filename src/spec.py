from abc import ABCMeta, abstractmethod
from typing import List, Dict
from string import ascii_uppercase
from openpyxl import Workbook, load_workbook
import pathlib
import os

from exceptions import DeviceModelDoesnotExistException, ParameterCannotBeNone


_Spec_Device_Info: List[Dict[str, str]] = []


class TestSpec(metaclass=ABCMeta):

    __DLink_Test_Time_Spec_File_Name: str = "D-Link-Wi-Fi-config_time_for_tool.xlsx"
    __Project_Root_Dir = str(pathlib.Path(__file__).absolute().parent.parent)
    __Testing_Time_Config_Path = os.path.join(__Project_Root_Dir, "cnf", "testing_time", __DLink_Test_Time_Spec_File_Name)

    # _Spec_Device_Info: List[Dict[str, str]] = None

    _DLink_Spec_WorkBook: Workbook = None
    _DLink_Spec_WorkBook_Sheet_Page: Workbook = None


    def __init__(self, read_only=True):
        self.loading_content(read_only=read_only)
        self.get_sheet_by_name(name=self.sheet_page_name)


    def loading_content(self, read_only=True) -> None:
        self._DLink_Spec_WorkBook = load_workbook(filename=self.__Testing_Time_Config_Path, read_only=read_only)


    def get_sheet_by_name(self, name: str) -> None:
        self._DLink_Spec_WorkBook_Sheet_Page = self._DLink_Spec_WorkBook.get_sheet_by_name(name=name)


    @property
    def get_device_info(self):
        return _Spec_Device_Info


    @property
    def current_sheet_page(self) -> Workbook:
        return self._DLink_Spec_WorkBook_Sheet_Page


    def _get_row_index(self, device_model: str, index: int) -> int:
        """
        Get the excel index value of target device module.
        """
        if device_model is not None:
            device_info = self.check_device_exist(device_model=device_model, get_info=True)
            return int(device_info["index"])
        elif index is not None:
            return index
        else:
            raise ParameterCannotBeNone


    def get_each_cases_testing_time(self, device_model: str, index: int = None) -> Dict[str, float]:
        """
        Get all test items and the time it needs to use.
        Return a dict type data, key is test item and value is time.
        """
        all_test_case_item = self.get_all_test_cases(device_model=device_model)

        if device_model is not None:
            has_device = self.check_device_exist(device_model=device_model)
            if has_device is True:
                each_case_testing_time = self.get_each_cases_testing_time_by_device(device_model=device_model)
            else:
                raise DeviceModelDoesnotExistException
        elif index is not None:
            each_case_testing_time = self.get_each_cases_testing_time_by_device(index=index)
        else:
            raise ParameterCannotBeNone

        return {test_item: test_time for test_item, test_time in zip(all_test_case_item, each_case_testing_time)}


    @abstractmethod
    def get_all_test_cases(self, device_model: str):
        pass


    @abstractmethod
    def get_each_cases_testing_time_by_device(self, device_model: str = None, index: int = None) -> List[str]:
        pass


    @property
    @abstractmethod
    def sheet_page_name(self) -> str:
        """
        The using sheet page name currently. 
        """
        pass


    @property
    @abstractmethod
    def first_datarow_index(self) -> int:
        """
        The index of first data row.
        """
        pass


    @property
    @abstractmethod
    def last_datarow_index(self) -> int:
        """
        The index of last data row.
        """
        pass


    @abstractmethod
    def get_all_device_model(self) -> List[str]:
        """
        Get all of device modules of current sheet page.
        """
        pass


    def __get_device_model_info(self):
        """
        Get device info about excel index, device series and device module.
        """
        device_model_info = self._get_all_row_data_with_columns(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                                columns="AB")
        for index, info in enumerate(device_model_info):
            __info = {"index": index+2, "device_series": info[0], "device_model": info[1].split(sep=f"_{self.sheet_page_name}")[0]}
            _Spec_Device_Info.append(__info)


    def check_device_exist(self, device_model: str, get_info: bool = False):
        """
        Check whether the device module exists or not.
        """
        if not _Spec_Device_Info:
            self.__get_device_model_info()

        for device_info in _Spec_Device_Info:
            if device_model == device_info["device_model"]:
                if get_info is True:
                    return device_info
                else:
                    return True
        raise DeviceModelDoesnotExistException


    def _get_ascii_char(self, start_ascii: str = 0, end_ascii: str = 26) -> str:
        """
        Get all the ascii characters between in 2 ascii.
        """
        start_ascii_index = ascii_uppercase.index(start_ascii)
        end_ascii_index = ascii_uppercase.index(end_ascii)
        return ascii_uppercase[start_ascii_index:end_ascii_index+1]


    def _get_all_row_data_with_row(self, sheet_page, row, columns) -> List[str]:
        """
        Get all value with multiple columns and one specific row.
        """
        result_data: List[str] = []
        for column in columns:
            cell_value = sheet_page[f"{column}{row}"].value  # the value of the specific cell
            result_data.append(cell_value)
        return result_data


    def _get_all_row_data_with_column(self, sheet_page, column, result_sheet=True) -> List[str]:
        """
        Get all data row with one specific column.
        """
        result_data: List[str] = []
        for row in range(self.first_datarow_index, self.last_datarow_index + 1):
            cell_value = sheet_page[f"{column}{row}"].value  # the value of the specific cell
            if cell_value is not None:
                if result_sheet:
                    result_data.append(cell_value)
                else:
                    result_data.append("_".join(cell_value.split(sep="_")[:-1]))
            else:
                result_data.append(cell_value)
        return result_data


    def _get_all_row_data_with_columns(self, sheet_page, columns) -> List[List[str]]:
        """
        Get all data row with multiple columns.
        """
        result_data: List[List[str]] = []
        for row in range(self.first_datarow_index, self.last_datarow_index + 1):
            data_row: List[str] = []
            for column in columns:
                cell_value = sheet_page[f"{column}{row}"].value  # the value of the specific cell
                data_row.append(cell_value)
            result_data.append(data_row)
        return result_data


    def close_workbook(self):
        """
        Close the WorkBook object.
        """
        self._DLink_Spec_WorkBook.close()




class ResultTestSpec(TestSpec):

    def get_all_device_model(self):
        return self._get_all_row_data_with_column(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                  column="C")


    @property
    def sheet_page_name(self):
        return "group test time"


    @property
    def first_datarow_index(self) -> int:
        return 5


    @property
    def last_datarow_index(self) -> int:
        return 63


    def get_all_test_cases(self, device_model: str):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=1,
                                               columns=self._get_ascii_char(start_ascii="D",
                                                                            end_ascii="J"))


    def get_each_cases_testing_time_by_device(self, device_model: str = None, index: int = None) -> List[str]:
        datarow_index = self._get_row_index(device_model=device_model, index=index)
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=datarow_index,
                                               columns=self._get_ascii_char(start_ascii="D",
                                                                            end_ascii="J"))


    def get_full_test(self):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=4,
                                               columns=self._get_ascii_char(start_ascii="D",
                                                                            end_ascii="J"))


    def get_regression_test(self):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=4,
                                               columns=self._get_ascii_char(start_ascii="K",
                                                                            end_ascii="N"))


    def get_easy_test(self):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=4,
                                               columns=self._get_ascii_char(start_ascii="O",
                                                                            end_ascii="O"))



class FullTestSpec(TestSpec):

    def get_all_device_model(self):
        return self._get_all_row_data_with_column(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                  column="B")


    @property
    def sheet_page_name(self):
        return "Full"


    @property
    def first_datarow_index(self) -> int:
        return 2


    @property
    def last_datarow_index(self) -> int:
        return 60


    def get_all_test_cases(self, device_model: str):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=1,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="R"))


    def get_each_cases_testing_time_by_device(self, device_model: str = None, index: int = None) -> List[str]:
        datarow_index = self._get_row_index(device_model=device_model, index=index)
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=datarow_index,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="R"))



class RegressionTestSpec(TestSpec):

    def get_all_device_model(self):
        return self._get_all_row_data_with_column(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                  column="B")


    @property
    def sheet_page_name(self):
        return "Regression"


    @property
    def first_datarow_index(self) -> int:
        return 2


    @property
    def last_datarow_index(self) -> int:
        return 61


    def get_all_test_cases(self, device_model: str):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=1,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="I"))


    def get_each_cases_testing_time_by_device(self, device_model: str = None, index: int = None) -> List[str]:
        datarow_index = self._get_row_index(device_model=device_model, index=index)
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=datarow_index,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="I"))



class EasyTestSpec(TestSpec):

    def get_all_device_model(self):
        return self._get_all_row_data_with_column(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                  column="B")


    @property
    def sheet_page_name(self):
        return "Easy"


    @property
    def first_datarow_index(self) -> int:
        return 2


    @property
    def last_datarow_index(self) -> int:
        return 60


    def get_all_test_cases(self, device_model: str):
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=1,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="D"))


    def get_each_cases_testing_time_by_device(self, device_model: str = None, index: int = None) -> List[str]:
        datarow_index = self._get_row_index(device_model=device_model, index=index)
        return self._get_all_row_data_with_row(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                               row=datarow_index,
                                               columns=self._get_ascii_char(start_ascii="C",
                                                                            end_ascii="D"))

