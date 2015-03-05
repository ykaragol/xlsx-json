package com.github.ykaragol.traveller;

import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SpreadSheetTraveller {

	private final List<SpreadSheetVisitor> visitors = new LinkedList<>();

	public void travel(Workbook wb) {
		visit(wb);
		int numberOfSheets = wb.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			Sheet sheet = wb.getSheetAt(i);
			visit(sheet);
			int firstRowNum = sheet.getFirstRowNum();
			int lastRowNum = sheet.getLastRowNum();
			for (int j = firstRowNum; j < lastRowNum; j++) {
				Row row = sheet.getRow(j);
				if (row == null) {
					continue;
				}
				visit(row);
				short firstCellNum = row.getFirstCellNum();
				short lastCellNum = row.getLastCellNum();
				for (int t = firstCellNum; t < lastCellNum; t++) {
					Cell cell = row.getCell(t);
					if (cell == null) {
						continue;
					}
					visit(cell);
				}
			}
		}
	}

	public void addVisitor(SpreadSheetVisitor visitor) {
		visitors.add(visitor);
	}

	public void removeVisitor(SpreadSheetVisitor visitor) {
		visitors.remove(visitor);
	}

	private void visit(Workbook wb) {
		for (SpreadSheetVisitor v : visitors) {
			v.visit(wb);
		}
	}

	private void visit(Sheet sheet) {
		for (SpreadSheetVisitor v : visitors) {
			v.visit(sheet);
		}
	}

	private void visit(Row row) {
		for (SpreadSheetVisitor v : visitors) {
			v.visit(row);
		}
	}

	private void visit(Cell cell) {
		for (SpreadSheetVisitor v : visitors) {
			v.visit(cell);
		}
	}

}
