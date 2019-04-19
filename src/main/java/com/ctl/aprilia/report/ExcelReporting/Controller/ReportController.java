package com.ctl.aprilia.report.ExcelReporting.Controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;

import com.ctl.aprilia.report.ExcelReporting.Service.ExcelService;

@Controller
public class ReportController {

	@Autowired
	ExcelService excelService;

	@RequestMapping(value = "/events/excel", method = RequestMethod.GET)
	public ResponseEntity<InputStreamResource> getEventsAsExcel(@RequestParam("to") String to, @RequestParam("from") String from) {
		try {
			Workbook workBook = excelService.getReport(to, from);
			File newFile = File.createTempFile("ApriliaReport", ".xls");
			FileOutputStream fileOutputStream = new FileOutputStream(newFile);
			workBook.write(fileOutputStream);
			InputStreamResource resource = new InputStreamResource(new FileInputStream(newFile));
			HttpHeaders headers = new HttpHeaders();
			headers.add("Content-Disposition", "attachment; filename=ApriliaReport.xls");
			return ResponseEntity.ok().headers(headers).body(resource);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
}
