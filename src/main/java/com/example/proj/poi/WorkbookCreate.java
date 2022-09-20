package com.example.proj.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookCreate {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("FinalSheet");

        Object sheetData[][] = { { "DN Date", "Material", "Purchase Order",
                "DN Number", "Location", "Truck No",
                "Gross Weight", "Tare Weight", "ACC Qty",
                "GR Number", "GR Date", "SUP DN No",
                "vendor", "Bag Weight", "DED Qty" } };

        int rowCount = 0;
        for (Object data[] : sheetData) {
            XSSFRow row = sheet.createRow(rowCount++);
            System.out.println(data);
            int columnCount = 0;
            for (Object value : data) {
                XSSFCell cell = row.createCell(columnCount++);
                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }

        String filePath = "sheetData.xlsx";
        FileOutputStream fileOutputStream = new FileOutputStream(filePath, true);
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        System.out.println("Data written");

    }

}
