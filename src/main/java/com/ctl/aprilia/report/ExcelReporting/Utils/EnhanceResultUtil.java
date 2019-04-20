package com.ctl.aprilia.report.ExcelReporting.Utils;

import java.util.Iterator;

import org.json.JSONObject;
import org.springframework.jdbc.core.JdbcTemplate;

import com.ctl.aprilia.report.ExcelReporting.Converter.ResultSetToJson;

public class EnhanceResultUtil {

	public static JSONObject enhanceResult(String query, JSONObject jsonObject, JdbcTemplate jdbcTemplate) {
		Iterator<String> keys = jsonObject.keys();
		String dataSet = "";
		while (keys.hasNext()) {
			dataSet = dataSet + "'"+keys.next()+"'" + ",";
		}
		dataSet = dataSet.substring(0, dataSet.length() - 1);
		try {
			JSONObject newDataSet = ResultSetToJson.convertResultSetIntoJSON(
					"select CUST_CODE, CUST_NAME,CUST_CITY from customer  where CUST_CODE in (" + dataSet + ")",
					jdbcTemplate);
			keys = newDataSet.keys();
			while (keys.hasNext()) {
				String key = keys.next();
				JSONObject parentDateSet = (JSONObject) jsonObject.get(key);
				JSONObject childDataSet = (JSONObject) newDataSet.get(key);
				Iterator<String> childKeys = childDataSet.keys();
				while (childKeys.hasNext()) {
					String currentKey = childKeys.next();
					parentDateSet.put(currentKey, childDataSet.get(currentKey));
				}
				jsonObject.put(key, parentDateSet);
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return jsonObject;
	}

}
