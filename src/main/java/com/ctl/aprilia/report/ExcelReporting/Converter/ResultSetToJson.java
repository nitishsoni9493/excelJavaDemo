package com.ctl.aprilia.report.ExcelReporting.Converter;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Locale;
import org.json.JSONObject;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.support.rowset.SqlRowSet;

public class ResultSetToJson {
	
	public static JSONObject convertResultSetIntoJSON(String sql, JdbcTemplate jdbcTemplate) throws Exception {
		JSONObject jsonArray = new JSONObject();
		SqlRowSet results = jdbcTemplate.queryForRowSet(sql);
		while (results.next()) {
			int total_rows = results.getMetaData().getColumnCount();
			JSONObject obj = new JSONObject();
			for (int i = 0; i < total_rows; i++) {
				String columnName = results.getMetaData().getColumnLabel(i + 1).toLowerCase();
				Object columnValue = results.getObject(i + 1);
				// if value in DB is null, then we set it to default value
				if (columnValue == null) {
					columnValue = "null";
				}
				/*
				 * Next if block is a hack. In case when in db we have values like price and
				 * price1 there's a bug in jdbc - both this names are getting stored as price in
				 * ResulSet. Therefore when we store second column value, we overwrite original
				 * value of price. To avoid that, i simply add 1 to be consistent with DB.
				 */
				if (obj.has(columnName)) {
					columnName += "1";
				}
				obj.put(columnName.toUpperCase(), columnValue);
			}
			jsonArray.put(obj.getString("CUST_CODE"),obj);
		}
		return jsonArray;
	}

	public static int converBooleanIntoInt(boolean bool) {
		if (bool)
			return 1;
		else
			return 0;
	}

	public static int convertBooleanStringIntoInt(String bool) {
		if (bool.equals("false"))
			return 0;
		else if (bool.equals("true"))
			return 1;
		else {
			throw new IllegalArgumentException("wrong value is passed to the method. Value is " + bool);
		}
	}

	public static double getDoubleOutOfString(String value, String format, Locale locale) {
		DecimalFormatSymbols otherSymbols = new DecimalFormatSymbols(locale);
		otherSymbols.setDecimalSeparator('.');
		DecimalFormat f = new DecimalFormat(format, otherSymbols);
		String formattedValue = f.format(Double.parseDouble(value));
		double number = Double.parseDouble(formattedValue);
		return Math.round(number * 100.0) / 100.0;
	}
}