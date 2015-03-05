package com.github.ykaragol.xlsxjson.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.bson.BSONObject;
import org.bson.BasicBSONObject;

import com.github.ykaragol.traveller.AbstractSpreadSheetVisitor;


public class JsonConverterVisitor extends AbstractSpreadSheetVisitor {
	
	private BasicBSONObject jsonObject= new BasicBSONObject();
	
	@Override
	public void visit(Cell cell) {
		int cellType = cell.getCellType();
		String cellValue = null;
		switch (cellType) {
		case Cell.CELL_TYPE_BLANK:
			cellValue = "";
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cellValue = cell.getBooleanCellValue() ? Boolean.TRUE.toString()
					: Boolean.FALSE.toString();
			break;
		case Cell.CELL_TYPE_ERROR:
			cellValue = "error!";
			break;
		case Cell.CELL_TYPE_FORMULA:
			cellValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				cellValue = cell.getDateCellValue().toString();
			} else {
				cellValue = String.valueOf(cell.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_STRING:
			cellValue = cell.getStringCellValue();
			break;
		default:
			throw new IllegalArgumentException("cell type invalid");
		}
		
		String sheetName = cell.getSheet().getSheetName();
		BSONObject jSheet = (BSONObject) jsonObject.get(sheetName);
		if(jSheet == null){
			jSheet = new BasicBSONObject();
			jsonObject.append(sheetName, jSheet);
		}		
		
		String rowIndex = String.valueOf(cell.getRowIndex());
		BSONObject jRow = (BSONObject) jSheet.get(rowIndex);
		if(jRow == null){
			jRow = new BasicBSONObject();
			jSheet.put(rowIndex, jRow);
		}
		
		String columnIndex = String.valueOf(cell.getColumnIndex());
		jRow.put(columnIndex, cellValue);
	}

	public BSONObject getJsonObject() {
		return jsonObject;
	}
	
}