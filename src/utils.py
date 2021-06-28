from abc import ABCMeta, abstractmethod
from typing import List, Dict
from string import ascii_uppercase
from openpyxl import Workbook, load_workbook



class SpecUtils:

    def get_all_device_model(self):
        return self._get_all_row_data_with_column(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                  column="B")


    def get_device_model_info(self):
        device_model_info = self._get_all_row_data_with_columns(sheet_page=self._DLink_Spec_WorkBook_Sheet_Page,
                                                                columns="AB")
        for index, info in enumerate(device_model_info):
            __info = {"index": index+2, "device_series": info[0], "device_model": info[1].split(sep=f"_{self.sheet_page_name}")[0]}
            _Spec_Device_Info.append(__info)


    def _get_all_row_data_with_column(self, sheet_page, column) -> List[str]:
        result_data: List[str] = []
        for row in range(self.first_datarow_index, self.last_datarow_index + 1):
            cell_value = sheet_page[f"{column}{row}"].value  # the value of the specific cell
            if cell_value is not None:
                result_data.append("_".join(cell_value.split(sep="_")[:-1]))
            else:
                result_data.append(cell_value)
        return result_data


    def _get_all_row_data_with_columns(self, sheet_page, columns) -> List[List[str]]:
        result_data: List[List[str]] = []
        for row in range(self.first_datarow_index, self.last_datarow_index + 1):
            data_row: List[str] = []
            for column in columns:
                cell_value = sheet_page[f"{column}{row}"].value  # the value of the specific cell
                data_row.append(cell_value)
            result_data.append(data_row)
        return result_data

