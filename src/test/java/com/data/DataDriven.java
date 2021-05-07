package com.data;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {
    public ArrayList<String> getData(String testCaseName) throws IOException {
        ArrayList<String> a = new ArrayList<>();
        FileInputStream fis = new FileInputStream("C:\\Users\\hamza\\Documents\\datademo.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        int sheet = workbook.getNumberOfSheets();
        for (int i = 0; i < sheet; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("demoData")) {
                XSSFSheet sheets = workbook.getSheetAt(i);
                Iterator<Row> row = sheets.iterator();
                Row firstRow = row.next();
                Iterator<Cell> ce = firstRow.cellIterator();
                int k = 0;
                int column = 0;
                while (ce.hasNext()) {
                    Cell value = ce.next();
                    if (value.getStringCellValue().equalsIgnoreCase(testCaseName)) {
                        column = k;

                    }
                    k++;
                }
                System.out.println(column);
                while ((row.hasNext())) {
                    Row r = row.next();
                    if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
                        Iterator<Cell> cv = r.cellIterator();
                        while (cv.hasNext()) {
                            Cell c = cv.next();
                            if (c.getCellTypeEnum()== CellType.STRING) {
                                a.add(c.getStringCellValue());
                            } else {
                                a.add(NumberToTextConverter.toText(c.getNumericCellValue()));

                            }


                        }
                    }
                }
            }
        }
        return a;
    }
}
