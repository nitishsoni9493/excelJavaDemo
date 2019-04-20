package com.ctl.aprilia.report.ExcelReporting.Converter;

import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.Gson;

public class JsonToExcelConverter {

	/*
	 * public static Workbook getExcelReport(JSONArray jsonArray) { Workbook
	 * workbook = new HSSFWorkbook(); // new HSSFWorkbook() for generating `.xls`
	 * file
	 * 
	 * CreationHelper createHelper = workbook.getCreationHelper();
	 * 
	 * // Create a Sheet Sheet sheet = workbook.createSheet("Report");
	 * 
	 * // Create a Font for styling header cells Font headerFont =
	 * workbook.createFont(); headerFont.setBold(true);
	 * headerFont.setFontHeightInPoints((short) 14);
	 * headerFont.setColor(IndexedColors.RED.getIndex());
	 * 
	 * // Create a CellStyle with the font CellStyle headerCellStyle =
	 * workbook.createCellStyle(); headerCellStyle.setFont(headerFont);
	 * 
	 * // Create Cell Style for formatting Date CellStyle dateCellStyle =
	 * workbook.createCellStyle();
	 * dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
	 * "dd-MM-yyyy"));
	 * 
	 * // Create Other rows and cells with employees data int rowNum = 1; for
	 * (Iterator<Object> iterator2 = jsonArray.iterator(); iterator2.hasNext();) {
	 * JSONObject type = (JSONObject) iterator2.next(); Row row =
	 * sheet.createRow(rowNum++); Iterator<String> keys = type.keys();
	 * 
	 * int j = 0; while (keys.hasNext()) { String key = keys.next(); if
	 * (type.get(key) instanceof JSONObject) {
	 * row.createCell(j).setCellValue(type.getString(key)); j++; } } } return
	 * workbook; }
	 */

	@SuppressWarnings("unchecked")
	public static Workbook getExcelReport(JSONObject jsonArray) {
		HSSFWorkbook workbook = new HSSFWorkbook();

		// Create a blank sheet
		HSSFSheet sheet = workbook.createSheet("Order Details");

		// This data needs to be written (Object[])
		Iterator<String> parentJson = jsonArray.keys();
		JSONObject newObject = new JSONObject();
		if (parentJson.hasNext()) {
			JSONObject object = (JSONObject) jsonArray.get(parentJson.next());
			Iterator<String> keys = object.keys();
			while (keys.hasNext()) {
				String key = keys.next();
				newObject.put(key, key.toUpperCase().replace("_", " "));
			}
		}
		int rownum = 0;
		rownum = addRow(sheet, newObject, rownum, workbook);
		ObjectMapper mapper = new ObjectMapper();
		Map<String, Object> map = new LinkedHashMap<String, Object>();
		try {
			map = mapper.readValue(jsonArray.toString(), LinkedHashMap.class);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		for (Map.Entry<String, Object> entry : map.entrySet()) {
			// this creates a new row in the sheet
			Gson gson = new Gson();
			String json = gson.toJson(entry.getValue(), LinkedHashMap.class);
			JSONObject objArr = new JSONObject(json);
			rownum = addRow(sheet, objArr, rownum, workbook);
		}
		return workbook;

	}

	private static Integer addRow(HSSFSheet sheet, JSONObject jsonObject, int rownum, HSSFWorkbook workbook) {
		HSSFRow row = sheet.createRow(rownum++);
		Iterator<String> valueSet = jsonObject.keys();
		int cellnum = 0;
		while (valueSet.hasNext()) {
			// this line creates a cell in the next column of that row
			String keyValue = valueSet.next();
			Object value = jsonObject.get(keyValue);
			HSSFCell cell = row.createCell(cellnum++);
			if (value instanceof String)
				cell.setCellValue((String) value);
			else if (keyValue.equalsIgnoreCase("CORECOMMENTS")) {
				HSSFCellStyle cs = workbook.createCellStyle();
				cs.setWrapText(true);
				cell.setCellStyle(cs);
				cell.setCellValue((String) value);
			} else if (value instanceof Object)
				cell.setCellValue((String) value.toString());
			int columnIndex = cell.getColumnIndex();
			sheet.autoSizeColumn(columnIndex);
		}
		return rownum;
	}

}
