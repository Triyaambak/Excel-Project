package com.demo.WriteInExcel;

import java.io.FileOutputStream;
//import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write {
	public static void write (int selections , String source , String entry , int start )
	{
		Cell cell;
		String source_name = null;
		int division = 0;		
		try {
			Workbook workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet("Write In Excel");
			String headername [] = {"Plot" , "Entry" , "Source" , "Remarks" , "Selections"};
			Row row = sheet.createRow(0);
			for(int i =0;i<headername.length;i++)
			{
				cell = row.createCell(i);
				cell.setCellValue(headername[i]);
			}
			//System.out.print(selections);
			//System.out.print("passed selection is "+selections);
			//System.out.print("passed source is "+source);
			//System.out.print("passed entry is "+entry);
			//System.out.print("Start is "+start);
			for(int i =1;i<=selections;i++)
			{				
				row = sheet.createRow(start++);	
				for(int j =0;j<3;j++)
				{
					if(j==0)
					{
						cell = row.createCell(0);
						cell.setCellValue(source);
					}
					if(j==1)
					{
						cell = row.createCell(1);
						cell.setCellValue(entry+"-"+i);
					}
					if(j==2)
					{
						cell = row.createCell(2);
						cell.setCellValue(source_name+":"+division+"-"+i);
					}					
				}
			}							
			FileOutputStream fileoutput = new FileOutputStream("/Java/Excel-Output.xlsx");
			workbook.write(fileoutput);
			fileoutput.close();
			workbook.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		//sc.close();
	}
}

