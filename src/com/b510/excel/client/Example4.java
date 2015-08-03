package com.b510.excel.client;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Example4 {
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("newworkbook.xls");
		Workbook wb = new HSSFWorkbook(fis); // or new
												// XSSFWorkbook("/somepath/test.xls")
		Sheet sheet = wb.getSheetAt(0);
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

		// suppose your formula is in B3
		CellReference cellReference = new CellReference("B3");
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());

		if (cell != null) {
			switch (evaluator.evaluateFormulaCell(cell)) {
			case Cell.CELL_TYPE_BOOLEAN:
				System.out.println(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				System.out.println(cell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				System.out.println(cell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_ERROR:
				System.out.println(cell.getErrorCellValue());
				break;

			// CELL_TYPE_FORMULA will never occur
			case Cell.CELL_TYPE_FORMULA:
				break;
			}
		}
	}
}
