from openpyxl.writer.excel import save_virtual_workbook
from datetime import datetime
from flask.wrappers import Request
from flask import Response

from .result import ResultReport



class ExportService:

    def export(self, request: Request):
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

