package com.demo.WriteInExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read extends Write {
	public static void main (String args[]) throws IOException
	{		
		int start =1;
		int selections = 0;
		String entry = null;
		String source = null;
		String filepath = "/Java/Excel-Input.xlsx";
		FileInputStream fileinput = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(filepath);
		XSSFSheet sheet = workbook.getSheet("Write In Excel");	
		for(int i = 1;i<=sheet.getLastRowNum();i++)
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
									   selections = (int)cell.getNumericCellValue();
										break;
						case STRING : System.out.print(cell.getStringCellValue());						
	 								  if(j==1)
										  entry = cell.getStringCellValue();
						 			  if(j==2)
										source = cell.getStringCellValue();
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
			//System.out.print("passed selection is "+selections);
			//System.out.print("passed source is "+source);
			//System.out.print("passed entry is "+entry);
			write(selections , source , entry , start);	
			start += selections;
			//System.out.print(start);
			System.out.println();
		}		
		fileinput.close();
		workbook.close();
	}

}
