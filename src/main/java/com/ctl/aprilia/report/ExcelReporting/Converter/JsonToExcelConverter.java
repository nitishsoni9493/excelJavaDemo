package com.ctl.aprilia.report.ExcelReporting.Converter;

import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;

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

	public static Workbook getExcelReport(JSONArray jsonArray) {
		HSSFWorkbook workbook = new HSSFWorkbook();

		// Create a blank sheet
		HSSFSheet sheet = workbook.createSheet("Order Details");

		// This data needs to be written (Object[])
		JSONObject jsonObject = (JSONObject) jsonArray.get(0);
		Iterator<String> keys = jsonObject.keys();
		JSONObject newObject = new JSONObject();
		while (keys.hasNext()) {
			String key = keys.next();
			newObject.put(key, key.toUpperCase().replace("_", " "));
		}

		jsonArray.put(0, newObject);
		keys = jsonObject.keys();

		int rownum = 0;
		for (int i = 0; i < jsonArray.length(); i++) {
			// this creates a new row in the sheet
			HSSFRow row = sheet.createRow(rownum++);
			JSONObject objArr = (JSONObject) jsonArray.get(i);
			Iterator<String> valueSet = objArr.keys();
			int cellnum = 0;
			while (valueSet.hasNext()) {
				// this line creates a cell in the next column of that row
				String keyValue = valueSet.next();
				Object value = objArr.get(keyValue);
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
			HSSFCell cell = row.createCell(cellnum++);
			HSSFCellStyle cs = workbook.createCellStyle();
			cs.setWrapText(true);
			cell.setCellStyle(cs);
			if (i == 0)
				cell.setCellValue("CORE COMMENTS");
			else
				cell.setCellValue(
						(String) "Pellentesque habitant morbi tristique \n senectus et netus et malesuada fames ac \n turpis egestas. Vestibulum tortor quam,\n feugiat vitae, ultricies eget,\n tempor sit amet, ante. Donec eu libero sit \n amet quam egestas semper.\n Aenean ultricies mi vitae est.\n Mauris placerat eleifend leo.\n Quisque sit amet est et sapien \n ullamcorper pharetra.\n Vestibulum erat wisi,\n condimentum sed, commodo vitae,\n ornare sit amet, wisi. Aenean fermentum,\n elit eget tincidunt condimentum,\n eros ipsum rutrum orci, sagittis\n tempus lacus enim ac dui.");
			int columnIndex = cell.getColumnIndex();
			sheet.autoSizeColumn(columnIndex);
		}
		return workbook;
	}

}
