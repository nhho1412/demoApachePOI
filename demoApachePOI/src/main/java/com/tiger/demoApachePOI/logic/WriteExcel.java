package com.tiger.demoApachePOI.logic;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tiger.demoApachePOI.model.Employee;

public class WriteExcel {
	public static final int COLUMN_INDEX_ID = 0;
	public static final int COLUMN_INDEX_FIRST_NAME = 1;
	public static final int COLUMN_INDEX_LAST_NAME = 2;
	public static final int COLUMN_INDEX_ADDRESS = 3;
	public static final int COLUMN_INDEX_SALARY = 4;
	public static final int COLUMN_INDEX_ALLOWANCE = 5;
	public static final int COLUMN_INDEX_TOTAL_MONEY = 6;
	public static int workbookType = 0;

	public static void main(String[] args) throws FileNotFoundException, IOException {
		List<Employee> employees = getEmployees();
		String fileExcelPath = "excelEx.xlsx";
		write(fileExcelPath, employees);
	}

	private static List<Employee> getEmployees() {
		List<Employee> employees = new ArrayList<Employee>();
		
		for (int i = 1; i <= 50; i++) {
			employees.add(getRandomEmployee(i));
		}

		return employees;
	}

	private static Employee getRandomEmployee(int i) {
		Random rd = new Random();
		
		Employee employee = new Employee();
		
		employee.setId(i);
		
		List<String> firstNameList = Arrays.asList("Hổ", "Huyền", "Hoa", "Huệ", "Cúc", "Lan", "Lộc", "Ngũ", "Tứ", "Tam", "Nhị");
		employee.setFirstName(firstNameList.get(rd.nextInt(firstNameList.size())));
		
		List<String> lastNameList = Arrays.asList("Nguyễn", "Trần", "Phạm", "Huỳnh", "Hoàng", "Lê", "Đào");
		List<String> middleNameList = Arrays.asList("Công", "Văn", "Huỳnh", "Nhã", "Như", "Ngọc", "Gia");
		employee.setLastName(lastNameList.get(rd.nextInt(lastNameList.size())) + " " + middleNameList.get(rd.nextInt(middleNameList.size())));
		
		List<String> addressList = Arrays.asList("Huế", "Đà Nẵng", "Hà Nội", "Quảng Nam", "Hồ Chí Minh", "Đồng Nai", "Vũng Tàu", "Hải Dương", "Quảng Ninh");
		employee.setAddress(addressList.get(rd.nextInt(addressList.size())));
		
		employee.setSalary(new BigDecimal(500000 * i));
		employee.setAllowance(new BigDecimal(100000 * i));
		employee.setTotalMoney(employee.getSalary().add(employee.getAllowance()));
		
		return employee;
	}
	
	public static void write(String fileExcelPath, List<Employee> employees) throws FileNotFoundException, IOException {
		Workbook workbook = getWorkbook(fileExcelPath);

		Sheet sheet = workbook.createSheet("Employees");

		int rowIndex = 0;
		writeHeader(sheet, rowIndex);

		for (Employee employee : employees) {
			rowIndex++;
			Row row = sheet.createRow(rowIndex);
			wirteContent(sheet, row, employee);
		}
		
		writeFooter(sheet, rowIndex + 1);

		autoSizeColumn(sheet);
		saveExcel(workbook, fileExcelPath);
		workbook.close();
	}

	private static void autoSizeColumn(Sheet sheet) {
		for (int columnIndex = 0; columnIndex <= COLUMN_INDEX_TOTAL_MONEY; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}
	}

	private static Workbook getWorkbook(String fileExcelPath) {
		Workbook workbook = null;
		if (fileExcelPath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();
			workbookType = 0;
		} else if (fileExcelPath.endsWith("xls")) {
			workbook = new HSSFWorkbook();
			workbookType = 1;
		} else {
			throw new IllegalAccessError("File " + fileExcelPath + " is not Excel file");
		}

		return workbook;
	}

	private static void writeFooter(Sheet sheet, int rowIndex) {
		CellStyle cellStyle = getCellStyleContentAndFooter(rowIndex, sheet);
		Row row = sheet.createRow(rowIndex);
		
		Cell cell = row.createCell(COLUMN_INDEX_ID);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(COLUMN_INDEX_FIRST_NAME);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(COLUMN_INDEX_LAST_NAME);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(COLUMN_INDEX_ADDRESS);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(COLUMN_INDEX_SALARY, CellType.FORMULA);
		cell.setCellStyle(cellStyle);
		cell.setCellFormula("SUM(E1:E"+ (rowIndex - 1) +")");

		cell = row.createCell(COLUMN_INDEX_ALLOWANCE, CellType.FORMULA);
		cell.setCellStyle(cellStyle);
		cell.setCellFormula("SUM(F1:F"+ (rowIndex - 1) +")");

		cell = row.createCell(COLUMN_INDEX_TOTAL_MONEY, CellType.FORMULA);
		cell.setCellStyle(cellStyle);
		cell.setCellFormula("SUM(G1:G"+ (rowIndex - 1) +")");
	}
	
	private static void writeHeader(Sheet sheet, int rowIndex) {

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Tahoma");
		font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		font.setColor(IndexedColors.YELLOW.getIndex());

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Row row = sheet.createRow(rowIndex);
		Cell cell = row.createCell(COLUMN_INDEX_ID);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("ID");

		cell = row.createCell(COLUMN_INDEX_FIRST_NAME);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("FIRST NAME");

		cell = row.createCell(COLUMN_INDEX_LAST_NAME);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("LAST NAME");

		cell = row.createCell(COLUMN_INDEX_ADDRESS);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("ADDRESS");

		cell = row.createCell(COLUMN_INDEX_SALARY);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("SALARY");

		cell = row.createCell(COLUMN_INDEX_ALLOWANCE);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("ALLOWANCE");

		cell = row.createCell(COLUMN_INDEX_TOTAL_MONEY);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("TOTAL MONEY");
	}

	private static CellStyle getCellStyleContentAndFooter(int rowNum, Sheet sheet) {
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Tahoma");
		font.setFontHeightInPoints((short) 12);

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);

		if (rowNum % 2 == 0) {
			cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}

		return cellStyle;
	}

	private static void wirteContent(Sheet sheet, Row row, Employee employee) {
		CellStyle cellStyle = getCellStyleContentAndFooter(row.getRowNum(), sheet);

		Cell cell = row.createCell(COLUMN_INDEX_ID);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(employee.getId());

		cell = row.createCell(COLUMN_INDEX_FIRST_NAME);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.STRING);
		cell.setCellValue(employee.getFirstName());

		cell = row.createCell(COLUMN_INDEX_LAST_NAME);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.STRING);
		cell.setCellValue(employee.getLastName());

		cell = row.createCell(COLUMN_INDEX_ADDRESS);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.STRING);
		cell.setCellValue(employee.getAddress());

		cell = row.createCell(COLUMN_INDEX_SALARY);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(employee.getSalary().doubleValue());

		cell = row.createCell(COLUMN_INDEX_ALLOWANCE);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(employee.getAllowance().doubleValue());

		cell = row.createCell(COLUMN_INDEX_TOTAL_MONEY);
		cell.setCellStyle(cellStyle);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(employee.getTotalMoney().doubleValue());
	}

	private static void saveExcel(Workbook workbook, String fileExcelPath) throws FileNotFoundException, IOException {
		try (OutputStream os = new FileOutputStream(fileExcelPath)) {
			workbook.write(os);
		}
	}
}
