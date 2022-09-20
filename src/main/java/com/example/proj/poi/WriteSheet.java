package com.example.proj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteSheet {
    public static void main(String[] args) throws IOException {
        String filePath = "SAMPLE Account 1.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook myWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet mySheet = myWorkbook.getSheetAt(3);
        XSSFSheet myOtherSheet = myWorkbook.getSheetAt(2);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("FinalSheet");
        int cellNumber = 0;

        Object sheetData[][] = { { "Date", "DC NO", "Truck No",
                "Material", "Party Name", "Place",
                "Rec", "ACC", "Price",
                "Amount", "Tax", "Paid 1",
                "Date 1", "Paid 2", "Date 2",
                "load", "DateL", "FRT", "T.paid", "Balance" } };

        int rowCount = 0;
        for (Object data[] : sheetData) {
            XSSFRow row = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Object value : data) {
                XSSFCell cell = row.createCell(columnCount++);
                cell.setCellValue((String) value);
            }
        }

        int readRows = mySheet.getLastRowNum();
        for (int i = 2; i < readRows; i++) {
            XSSFRow myRow = mySheet.getRow(i);
            XSSFRow row = sheet.createRow(rowCount++);
            cellNumber = 0;
            XSSFCell cell = myRow.getCell(cellNumber);
            int writeCellNumber = 0;
            Object cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
            cellNumber = 5;
            cell = myRow.getCell(cellNumber);
            writeCellNumber = 2;
            cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
            cellNumber = 1;
            cell = myRow.getCell(cellNumber);
            writeCellNumber = 3;
            cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
        }

        readRows = myOtherSheet.getLastRowNum();
        for (int i = 2; i < readRows; i++) {
            XSSFRow myRow = myOtherSheet.getRow(i);
            XSSFRow row = sheet.createRow(rowCount++);
            cellNumber = 0;
            XSSFCell cell = myRow.getCell(cellNumber);
            int writeCellNumber = 0;
            Object cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
            cellNumber = 4;
            cell = myRow.getCell(cellNumber);
            writeCellNumber = 2;
            cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
            cellNumber = 1;
            cell = myRow.getCell(cellNumber);
            writeCellNumber = 3;
            cellValue = getCellValue(cell);
            writeCell(writeCellNumber, row, cellValue);
        }

        String writeFile = "sheetData.xlsx";
        FileOutputStream fileOutputStream = new FileOutputStream(writeFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Data written");
    }

    private static void writeCell(int cellNumber, XSSFRow row, Object cellValue) {
        if (cellValue instanceof String)
            row.createCell(cellNumber).setCellValue((String) cellValue);
        if (cellValue instanceof Integer)
            row.createCell(cellNumber).setCellValue((Integer) cellValue);
        if (cellValue instanceof Boolean)
            row.createCell(cellNumber).setCellValue((Boolean) cellValue);
    }

    private static Object getCellValue(XSSFCell cell) {
        switch (cell.getCellType()) {
            case STRING:
                // System.out.println(cell.getStringCellValue());
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            default:
                break;
        }
        return cell.getRawValue();
    }
}
