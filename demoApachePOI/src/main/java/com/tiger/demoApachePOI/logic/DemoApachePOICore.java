package com.tiger.demoApachePOI.logic;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class DemoApachePOICore {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		DemoApachePOICore demoApachePOILogic = new DemoApachePOICore();
		
		demoApachePOILogic.workbook2003();
	}
	
	public void workbook2003() throws FileNotFoundException, IOException {
		
		/**
		 * workbook
		 */
		// create workbook
		Workbook wb03 = new HSSFWorkbook();
		String excelFilePath = "books2003.xls";
		
		// Open an existing workbook
		// Way 1
		// HSSFWorkbook wb03 = new HSSFWorkbook(new FileInputStream(new File(excelFilePath)));
		// Way 2
		// Workbook wb03 = WorkbookFactory.create(new File(excelFilePath));
		
		/**
		 * Sheet
		 */
		// create default name
		// Sheet sheet = wb03.createSheet();
		// create another name
		String safeName = WorkbookUtil.createSafeSheetName("sheetAnotherName");
		Sheet sheetAnotherName = wb03.createSheet(safeName);
		
		// get sheet to workbook
		// Sheet sheet1 = wb03.getSheetAt(0);
		// Sheet sheet2 = wb03.getSheet("sheetAnotherName");
		
		/**
		 * Row
		 */
		// create row
		Row row = sheetAnotherName.createRow(0);
		
		// get row
		// Row row = sheetAnotherName.getRow(0);
		
		/**
		 * Cell
		 */
		// create cell
		Cell cell = row.createCell(0);
		
		// get cell
		// Cell cell1 = row.getCell(0);
		
		// cellType
		cell = row.createCell(0, CellType.FORMULA);
		
		cell.setCellValue("abc");
		
		createOutputFile(wb03, excelFilePath);
		System.out.println("Done");
	}
	
    private static void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
        try (OutputStream os = new FileOutputStream(excelFilePath)) {
            workbook.write(os);
        }
    }
}
