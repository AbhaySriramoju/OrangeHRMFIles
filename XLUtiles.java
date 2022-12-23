package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtiles {
	
	public static FileInputStream filein;
	public static FileOutputStream fileOut;
	public static Workbook Wbook;
	public static Sheet Sheet;
	public static Row row;
	public static Cell cell;
	public static CellStyle style;
	
	
	public static int getRowCount(String xlFile, String sheetName) throws IOException 
	{
		
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		int RowCount = Sheet.getLastRowNum();
		Wbook.close();
		filein.close();
		return RowCount;

	}
	
	public static short getColCount(String xlFile, String sheetName, int rowNum) throws IOException
	{	
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		Row row = Sheet.getRow(rowNum);
		short colCount = row.getLastCellNum();
		Wbook.close();
		filein.close();
		return colCount;
	
	}	
	
	public static String getStringCellData(String xlFile, String sheetName, int rowNum, int cell) throws IOException
	{
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		Row row = Sheet.getRow(rowNum);
		Cell cells = row.getCell(cell);
		String cellData = cells.getStringCellValue();
		Wbook.close();
		filein.close();
		return cellData;
		
		
	}
	
	public static double getNumericCellData(String xlFile, String sheetName, int rowNum, int cell) throws IOException
	{
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		Row row = Sheet.getRow(rowNum);
		Cell cells = row.getCell(cell);
		double cellData = cells.getNumericCellValue();
		Wbook.close();
		filein.close();
		return cellData;
		
	}
	
//	public static boolean getBooleanCellData(String xlFile, String sheetName, int rows, int cell) throws IOException
//	{
//
//		filein = new FileInputStream(xlFile);
//		Wbook = new XSSFWorkbook(filein);
//		Sheet = Wbook.getSheet(sheetName);
//		Row row = Sheet.getRow(rows);
//		Cell cells = row.getCell(cell);
//		boolean cellData = cells.getBooleanCellValue();
//		Wbook.close();
//		filein.close();
//		return cellData;
//		
//	
//		
//	}
	
	public static void setCellData(String xlFile, String sheetName, int rowNum, int cellNum, String Data) throws IOException
	{	
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		row = Sheet.getRow(rowNum);
		Cell cellcreate = row.createCell(cellNum);
		cellcreate.setCellValue(Data);
		
		fileOut = new FileOutputStream(xlFile);
		Wbook.write(fileOut);
		Wbook.close();
		
		
	}	
	
	public static void fillGreenColor(String xlFile, String sheetName, int rowNum, int cellNum) throws IOException
	{	
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		row = Sheet.getRow(rowNum);
		cell = row.getCell(cellNum);
		
		
		style = Wbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		
		fileOut = new FileOutputStream(xlFile);
		Wbook.write(fileOut);
		Wbook.close();
	}	
	
	public static void fillRedColor(String xlFile, String sheetName, int rowNum, int cellNum) throws IOException
	{
		filein = new FileInputStream(xlFile);
		Wbook = new XSSFWorkbook(filein);
		Sheet = Wbook.getSheet(sheetName);
		row = Sheet.getRow(rowNum);
		cell = row.getCell(cellNum);
		
		
		style = Wbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		
		fileOut = new FileOutputStream(xlFile);
		Wbook.write(fileOut);
		Wbook.close();
		
		
	}
	
	
	
	
	
	
	
	
}























