package com.tiger.demoApachePOI.logic;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tiger.demoApachePOI.model.Employee;

public class ReadExcel {
	public static final int COLUMN_INDEX_ID = 0;
	public static final int COLUMN_INDEX_FIRST_NAME = 1;
	public static final int COLUMN_INDEX_LAST_NAME = 2;
	public static final int COLUMN_INDEX_ADDRESS = 3;
	public static final int COLUMN_INDEX_SALARY = 4;
	public static final int COLUMN_INDEX_ALLOWANCE = 5;
	public static final int COLUMN_INDEX_TOTAL_MONEY = 6;

	public static void main(String[] args) throws IOException {
		String fileExcelPath = "excelEx.xlsx";
		List<Employee> employees = getEmployeesToExcel(fileExcelPath);

		for (Employee employee : employees) {
			System.out.println(employee);
		}
		
		System.out.println("done");
	}

	private static List<Employee> getEmployeesToExcel(String fileExcelPath) throws IOException {
		List<Employee> employees = new ArrayList<Employee>();

		InputStream inputStream = new FileInputStream(new File(fileExcelPath));
		Workbook workbook = getWorkbook(inputStream, fileExcelPath);

		Sheet sheet = workbook.getSheetAt(0);

		Iterator<Row> iterator = sheet.iterator();
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			if (nextRow.getRowNum() == 0) {
				continue;
			}

			Iterator<Cell> cellIterator = nextRow.cellIterator();
			Employee employee = new Employee();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				Object cellValue = getCellValue(cell);
				if (cellValue == null || cellValue.toString().isEmpty()) {
					continue;
				}

				int columnIndex = cell.getColumnIndex();
				switch (columnIndex) {
				case COLUMN_INDEX_ID:
					employee.setId(new BigDecimal((double) cellValue).intValue());
					break;
				case COLUMN_INDEX_FIRST_NAME:
					employee.setFirstName((String) cellValue);
					break;
				case COLUMN_INDEX_LAST_NAME:
					employee.setLastName((String) cellValue);
					break;
				case COLUMN_INDEX_ADDRESS:
					employee.setAddress((String) cellValue);
					break;
				case COLUMN_INDEX_ALLOWANCE:
					employee.setAllowance(new BigDecimal((double) cellValue));
					break;
				case COLUMN_INDEX_SALARY:
					employee.setSalary(new BigDecimal((double) cellValue));
					break;
				case COLUMN_INDEX_TOTAL_MONEY:
					employee.setTotalMoney(new BigDecimal((double) cellValue));
					break;
				default:
					break;
				}
			}
			employees.add(employee);
		}
		workbook.close();
		inputStream.close();

		return employees;
	}

	private static Object getCellValue(Cell cell) {
		CellType cellType = cell.getCellTypeEnum();
		Object cellValue = null;

		switch (cellType) {
		case BOOLEAN:
			cellValue = cell.getBooleanCellValue();
			break;
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case NUMERIC:
			cellValue = cell.getNumericCellValue();
			break;
		case FORMULA:
			Workbook workbook = cell.getSheet().getWorkbook();
			FormulaEvaluator fe = workbook.getCreationHelper().createFormulaEvaluator();
			cellValue = fe.evaluate(cell).getNumberValue();
			break;
		case BLANK:
		case ERROR:
			break;
		default:
			break;
		}

		return cellValue;
	}

	private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
		Workbook workbook = null;
		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook(inputStream);
		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook(inputStream);
		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}

		return workbook;
	}
}
