package com.example.proj.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
    public static void main(String[] args) throws IOException {
        String filePath = "Sample_Student_data_xlsx.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

        // XSSFSheet sheet= workbook.getSheet("Sheet1");
        XSSFSheet sheet = workbook.getSheetAt(0);

        // int rows = sheet.getLastRowNum();
        // int columns = sheet.getRow(1).getLastCellNum();

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()) {
            XSSFRow row = (XSSFRow) iterator.next();
            Iterator cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;
                    default:
                        System.out.println(cell.getRawValue());
                        break;
                }
                System.out.print(" | ");
            }
            System.out.println();
        }

    }
}
