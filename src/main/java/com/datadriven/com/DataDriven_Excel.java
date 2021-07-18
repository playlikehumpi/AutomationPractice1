package com.datadriven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;
import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class DataDriven_Excel {

	public static void particular_Data() throws IOException {

		File excelFile = new File("F:\\Eclipse Workarea\\AutomationPractice1\\logindata.xlsx");
		FileInputStream fileStream = new FileInputStream(excelFile);
		Workbook excelBook = new XSSFWorkbook(fileStream);

		org.apache.poi.ss.usermodel.Sheet workSheet = excelBook.getSheetAt(0);
		Row row = workSheet.getRow(1);
		org.apache.poi.ss.usermodel.Cell cellData = row.getCell(1);
		org.apache.poi.ss.usermodel.CellType cellTYpe = cellData.getCellType();

		if (cellTYpe.equals(cellTYpe.STRING)) {
			String stringCellValue = cellData.getStringCellValue();
			System.out.println(stringCellValue);
		}

		else if (cellTYpe.equals(cellTYpe.NUMERIC)) {
			double numericCellValue = cellData.getNumericCellValue();
			int value = (int) numericCellValue; // narrow casting
			System.out.println(value);
		}

	}

	public static void GetAllData() throws IOException {

		File excelFile = new File("F:\\Eclipse Workarea\\AutomationPractice1\\logindata.xlsx");

		FileInputStream fis = new FileInputStream(excelFile);
		Workbook wbook = new XSSFWorkbook(fis);

		org.apache.poi.ss.usermodel.Sheet sheet = wbook.getSheetAt(1);

		int rowsize = sheet.getPhysicalNumberOfRows();
		
		System.out.println("Row Data   " + rowsize);
		 
		
		for (int i = 0; i < rowsize; i++) {
			System.out.println("Column Data");
			
			for (int j = 0; j < 4; j++) {
				
				
				Row row= sheet.getRow(j);
				
				org.apache.poi.ss.usermodel.Cell cell=row.getCell(i);
				System.out.println(cell);
				
				
			}
			
		}
		
		
		
//		for (int i = 0; i < rowsize; i++) {
//
//			Row row = sheet.getRow(i);
//			int cellsize = row.getPhysicalNumberOfCells();
//			for (int j = 0; j < cellsize; j++) {
//				org.apache.poi.ss.usermodel.Cell celldata = row.getCell(j);
//				org.apache.poi.ss.usermodel.CellType cellType = celldata.getCellType();
//
//				if (cellType.equals(cellType.STRING)) {
//					String stringCellValue = celldata.getStringCellValue();
//					System.out.println(stringCellValue);
//				}
//
//				else if (cellType.equals(cellType.NUMERIC)) {
//					double numericCellValue = celldata.getNumericCellValue();
//					int value = (int) numericCellValue; // narrow casting
//					System.out.println(value);
//				}
//
//			}
//
//		}

	}

	public static void writeData() throws IOException {

		File excelFile = new File("F:\\Eclipse Workarea\\AutomationPractice1\\logindata.xlsx");
		FileInputStream fileStream = new FileInputStream(excelFile);
		Workbook excelBook = new XSSFWorkbook(fileStream);

		org.apache.poi.ss.usermodel.Sheet createSheet = excelBook.createSheet("StudentDetails");
		Row createRow = createSheet.createRow(0);
		org.apache.poi.ss.usermodel.Cell createCell = createRow.createCell(0);
		
		createCell.setCellValue("Student Name");
		excelBook.getSheet("StudentDetails").getRow(0).createCell(1).setCellValue("RollNo");
		
		excelBook.getSheet("StudentDetails").createRow(1).createCell(0).setCellValue("Judit");
		excelBook.getSheet("StudentDetails").getRow(1).createCell(1).setCellValue("45697");
		
		
		excelBook.getSheet("StudentDetails").createRow(2).createCell(0).setCellValue("Susan");
		excelBook.getSheet("StudentDetails").getRow(2).createCell(1).setCellValue("867532");
		
		excelBook.getSheet("StudentDetails").createRow(3).createCell(0).setCellValue("Sofia");
		excelBook.getSheet("StudentDetails").getRow(3).createCell(1).setCellValue("98765");
	
	
		FileOutputStream fos = new FileOutputStream(excelFile);
		excelBook.write(fos);
		excelBook.close();
		System.out.println("Sheet created successfully");

	}

	public static void main(String[] args) throws IOException {

		particular_Data();

	    GetAllData();
		//writeData();
	}

}
