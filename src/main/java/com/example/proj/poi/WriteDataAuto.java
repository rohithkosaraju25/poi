package com.example.proj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataAuto {
    private static HashMap<String, Integer> header = new HashMap<>();

    public static void main(String[] args) throws IOException {
        dataSetup();
        String filePath = "SAMPLE Account 1.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook myWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet mySheet = myWorkbook.getSheetAt(3);
        XSSFSheet myOtherSheet = myWorkbook.getSheetAt(2);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("FinalSheet");

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
        int readColumn = mySheet.getRow(0).getLastCellNum();
        for (int i = 2; i < readRows; i++) {
            XSSFRow row = mySheet.getRow(i);
            XSSFRow writeRow = sheet.createRow(rowCount++);
            for (int c = 0; c < readColumn; c++) {
                XSSFCell cell = row.getCell(c);
                String headerName = mySheet.getRow(1).getCell(c).getStringCellValue().toUpperCase();
                if (header.containsKey(headerName)) {
                    Object cellValue = getCellValue(cell);
                    writeCell(header.get(headerName), writeRow, cellValue);
                }
            }
        }

        readRows = myOtherSheet.getLastRowNum();
        readColumn = myOtherSheet.getRow(0).getLastCellNum();
        for (int i = 2; i < readRows; i++) {
            XSSFRow row = myOtherSheet.getRow(i);
            XSSFRow writeRow = sheet.createRow(rowCount++);
            for (int c = 0; c < readColumn; c++) {
                XSSFCell cell = row.getCell(c);
                String headerName = myOtherSheet.getRow(0).getCell(c).getStringCellValue().toUpperCase();
                if (header.containsKey(headerName)) {
                    Object cellValue = getCellValue(cell);
                    writeCell(header.get(headerName), writeRow, cellValue);
                }
            }
        }

        String writeFile = "sheetDataAuto.xlsx";
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

    private static void dataSetup() {
        header.put("DATE", 0);
        header.put("DN DATE", 0);
        header.put("DC NO", 1);
        header.put("TRUCK NO", 2);
        header.put("MATERIAL", 3);
        header.put("PARTY NAME", 4);
        header.put("PLACE", 5);
        header.put("REC.", 6);
        header.put("ACC.", 7);
        header.put("PRICE", 8);
        header.put("AMOUNT", 9);
        header.put("TAX", 10);
        header.put("PAID 1", 11);
        header.put("DATE 1", 12);
        header.put("PAID 2", 13);
        header.put("DATE 2", 14);
        header.put("LOAD", 15);
        header.put("DATE L", 16);
        header.put("FRT", 17);
        header.put("F DATE", 18);
        header.put("DATE 2", 19);
        header.put("T.PAID", 20);
        header.put("BALANCE", 21);
    }
}
