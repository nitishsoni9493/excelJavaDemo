package com.ctl.aprilia.report.ExcelReporting.Service;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelService {

	Workbook getReport(String to, String from);

}
