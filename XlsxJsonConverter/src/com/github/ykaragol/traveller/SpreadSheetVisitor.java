package com.github.ykaragol.traveller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface SpreadSheetVisitor {

	public abstract void visit(Workbook workbook);

	public abstract void visit(Sheet sheet);

	public abstract void visit(Row row);

	public abstract void visit(Cell cell);

}
