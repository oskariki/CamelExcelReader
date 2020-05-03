package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelReaderBean {
    public ArrayList readExcelCellData(InputStream body) throws Exception {

        String name = "";
        int phonenumber;

        ArrayList<String> cell_values = new ArrayList<String>();

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(body);
            XSSFSheet sheet = workbook.getSheetAt(0);
            boolean headersFound = false;
            int colNum;
            for (Iterator rit = sheet.rowIterator(); rit.hasNext(); ) {
                XSSFRow row = (XSSFRow) rit.next();
                if (!headersFound) {  // Skip the first row with column headers
                    headersFound = true;
                    continue;
                }
                colNum = 0;
                for (Iterator cit = row.cellIterator(); cit.hasNext(); ++colNum) {
                    XSSFCell cell = (XSSFCell) cit.next();
                    if (headersFound)
                        switch (colNum) {
                            case 0: // Name
                                name = cell.getRichStringCellValue().getString();
                                cell_values.add(name);
                                break;
                            case 1: // Phonenumber
                                phonenumber = (int)cell.getNumericCellValue();
                                cell_values.add(String.valueOf(phonenumber));
                                break;
                        }
                }
            }
            return cell_values;
        } catch (Exception e) {
            throw new RuntimeException("Unable to import Excel data", e);
        }
    }
}
