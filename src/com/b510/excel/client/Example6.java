package com.b510.excel.client;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Example6 {
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("newworkbook.xls");
		Workbook wb = new HSSFWorkbook(fis); // or new
												// XSSFWorkbook("/somepath/test.xls")
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
		for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
			Sheet sheet = wb.getSheetAt(sheetNum);
			for (Row r : sheet) {
				for (Cell c : r) {
					if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
						evaluator.evaluateFormulaCell(c);
					}
				}
			}
		}
	}
}
