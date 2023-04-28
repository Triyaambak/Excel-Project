package com.demo.WriteInExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write {
    public static void write(int selections, String source, String entry, int start) {
        Cell cell;
        String source_name = null;
        int division = 0;
        String filepath = "/Java/Excel-Output.xlsx";
        try {
            FileInputStream fileinput = new FileInputStream(filepath);
            Workbook workbook = new XSSFWorkbook(fileinput);
            Sheet sheet = workbook.getSheet("Write In Excel");
            int lastRowNum = sheet.getLastRowNum();
            String headername[] = { "Plot", "Entry", "Source", "Remarks", "Selections" };
            if (lastRowNum == 0) {
                // create header row if it doesn't exist
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < headername.length; i++) {
                    cell = headerRow.createCell(i);
                    cell.setCellValue(headername[i]);
                }
                lastRowNum++; // increment to start at data row
            }
            for (int i = 1; i <= selections; i++) {
                Row row = sheet.createRow(lastRowNum++);
                for (int j = 0; j < 3; j++) {
                    if (j == 0) {
                        cell = row.createCell(0);
                        cell.setCellValue(source);
                    }
                    if (j == 1) {
                        cell = row.createCell(1);
                        cell.setCellValue(entry + "-" + i);
                    }
                    if (j == 2) {
                        cell = row.createCell(2);
                        cell.setCellValue(source_name + ":" + division + "-" + i);
                    }
                }
            }
            FileOutputStream fileoutput = new FileOutputStream(filepath);
            workbook.write(fileoutput);
            fileoutput.close();
            workbook.close();
            fileinput.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
