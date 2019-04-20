package com.ctl.aprilia.report.ExcelReporting.Service;

import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import com.ctl.aprilia.report.ExcelReporting.Utils.EnhanceResultUtil;
import com.ctl.aprilia.report.ExcelReporting.Converter.JsonToExcelConverter;
import com.ctl.aprilia.report.ExcelReporting.Converter.ResultSetToJson;

@Service
public class ExcelServiceIMPL implements ExcelService {

	
	@Autowired
	JdbcTemplate jdbcTemplate;
	@Override
	public Workbook getReport(String to, String from) {
		String sql = "select * from daysorder";
		try {
			JSONObject jsonArray = ResultSetToJson.convertResultSetIntoJSON(sql, jdbcTemplate);
			EnhanceResultUtil.enhanceResult("select CUST_CODE, CUST_NAME from customer", jsonArray, jdbcTemplate);
			return JsonToExcelConverter.getExcelReport(jsonArray);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

}
