package com.obsqura1.automationmaventraining;

import java.io.FileInputStream;//import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
public class ExcelRead 
{
	static FileInputStream fis;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	public static String readStringData(int row, int column) throws Exception
	{
		fis = new FileInputStream("D:\\Java\\Maven_Excel.xlsx");
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet("First");
		XSSFRow r = sheet.getRow(row);
		XSSFCell c = r.getCell(column);
		return c.getStringCellValue();
	}
	public static double readNumericData (int row, int column) throws Exception
	{
		fis = new FileInputStream("D:\\Java\\Maven_Excel.xlsx");
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet ("First");
		XSSFRow r = sheet.getRow(row);
		XSSFCell c = r.getCell(column);
		return c.getNumericCellValue();
	}
	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		System.out.println(ExcelRead.readStringData(0,0));
		double d = ExcelRead.readNumericData(1, 1);
		System.out.println(d);
	}

}