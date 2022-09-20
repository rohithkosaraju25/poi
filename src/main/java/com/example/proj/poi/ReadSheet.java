package com.example.proj.poi;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadSheet {
    public static void main(String[] args) throws IOException {
        String filePath = "SAMPLE Account 1.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(3);

        int rows = sheet.getLastRowNum();
        int columns = sheet.getRow(1).getLastCellNum();

        /*
         * for (int r = 0; r < rows; r++) {
         * XSSFRow row = sheet.getRow(r);
         * for (int c = 0; c < columns; c++) {
         * XSSFCell cell = row.getCell(c);
         * getCells(cell);
         * System.out.print(" | ");
         * }
         * }
         */
        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(1);
            getCells(cell);
        }
        System.out.println();
    }

    private static void getCells(XSSFCell cell) {
        switch (cell.getCellType()) {
            case STRING:
                System.out.println(cell.getStringCellValue());
                break;
            case NUMERIC:
                System.out.print(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            default:
                // System.out.print(cell.getRawValue());
                break;
        }
    }
}
