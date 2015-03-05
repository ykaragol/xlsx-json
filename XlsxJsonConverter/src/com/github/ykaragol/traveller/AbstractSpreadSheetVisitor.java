package com.github.ykaragol.traveller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class AbstractSpreadSheetVisitor implements SpreadSheetVisitor {

	@Override
	public void visit(Cell cell) {
		;
	}

	@Override
	public void visit(Row row) {
		;
	}

	@Override
	public void visit(Sheet sheet) {
		;
	}

	@Override
	public void visit(Workbook workbook) {
		;
	}

}
