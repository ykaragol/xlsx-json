package com.github.ykaragol.xlsxjson;

import org.apache.poi.ss.usermodel.Workbook;
import org.bson.BSONObject;

import com.github.ykaragol.ssutils.WorkbookUtils;
import com.github.ykaragol.traveller.SpreadSheetTraveller;
import com.github.ykaragol.xlsxjson.converter.JsonConverterVisitor;


public class XlsxBsonConverter {

	public String convert(String fileName){
		Workbook wb = createWorkbook(fileName);
		if(wb == null)
			return null;
		BSONObject jsonObject = getBsonObject(wb);
		String json = getString(jsonObject);
		return json;
	}

	public Workbook createWorkbook(String fileName) {
		return new WorkbookUtils().createWorkbook(fileName);
	}
	
	public BSONObject getBsonObject(Workbook wb) {
		SpreadSheetTraveller t = new SpreadSheetTraveller();
		JsonConverterVisitor visitor = new JsonConverterVisitor();
		t.addVisitor(visitor);
		t.travel(wb);
		BSONObject jsonObject = visitor.getJsonObject();
		return jsonObject;
	}
	
	public String getString(BSONObject jsonObject) {
		return jsonObject.toString();
	}	

}
