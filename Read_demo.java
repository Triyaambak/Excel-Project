package com.demo.WriteInExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_demo {
	public static void main (String args[]) throws IOException
	{
		String filepath = "/Java/Excel-Input.xlsx";
		FileInputStream fileinput = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(filepath);
		XSSFSheet sheet = workbook.getSheet("Write In Excel");	
		for(int i = 0;i<sheet.getLastRowNum();i++)
		{
			XSSFRow row = sheet.getRow(i);
			int last_column = row.getLastCellNum();
			for(int j =0;j<last_column;j++)
			{
				try 
				{
					XSSFCell cell = row.getCell(j);
					switch(cell.getCellType())
					{
						case NUMERIC : System.out.print(cell.getNumericCellValue());
										break;
						case STRING : System.out.print(cell.getStringCellValue());
										break;
						case BOOLEAN : System.out.print(cell.getBooleanCellValue());
										break;
						default : break;
					}
				}
				catch (Exception e)
				{
					System.out.print("Null");
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		fileinput.close();
		workbook.close();
	}
}
