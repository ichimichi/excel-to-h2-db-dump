package org.example;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelUtil {

    public String path;
    public FileInputStream fileInputStream = null;
    public FileOutputStream fileOutputStream = null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private XSSFRow row = null;
    private XSSFCell cell = null;
    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

    public ExcelUtil(String path) {

        this.path = path;
        try {
            fileInputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheetAt(0);
            fileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void printStats() {
        System.out.printf("SHEET\t\tROWS\tCOLUMNS\n");
        for (String sheet : getSheetNames()) {
            System.out.printf("%s\t\t%s\t%s\n", sheet, getRowCount(sheet), getColumnCount(sheet));
        }
    }

    public void printSheet(String sheet) {
        int row = getRowCount(sheet);
        int col = getColumnCount(sheet);
        for (int i = 0; i < row; i++) {
            for (int j = 0; j < col; j++) {
                System.out.printf("%s\t\t", getCellData(sheet, i, j));
            }
            System.out.println();
        }
    }


    public boolean isSheetExist(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1) {
            index = workbook.getSheetIndex(sheetName.toUpperCase());
            if (index == -1)
                return false;
            else
                return true;
        } else
            return true;
    }

    public List<String> getSheetNames() {
        List<String> sheetList = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheetList.add(workbook.getSheetAt(i).getSheetName());
        }

        return sheetList;
    }

    public int getRowCount(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1)
            return 0;
        else {
            sheet = workbook.getSheetAt(index);
            int number = sheet.getLastRowNum() + 1;
            return number;
        }

    }

    public int getColumnCount(String sheetName) {
        // check if sheet exists
        if (!isSheetExist(sheetName))
            return -1;

        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(0);

        if (row == null)
            return -1;

        return row.getLastCellNum();

    }


    public String getCellData(String sheetName, int rowNum, String columnName) {
        try {
            if (rowNum <= 0)
                return "";

            int index = workbook.getSheetIndex(sheetName);
            int col_Num = -1;
            if (index == -1)
                return "";

            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                // System.out.println(row.getCell(i).getStringCellValue().trim());
                if (row.getCell(i).getStringCellValue().trim().equals(columnName.trim()))
                    col_Num = i;
            }
            if (col_Num == -1)
                return "";

            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(rowNum - 1);
            if (row == null)
                return "";
            cell = row.getCell(col_Num);

            if (cell == null)
                return "";

            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case FORMULA:
                case NUMERIC:
                    String cellText = String.valueOf(cell.getNumericCellValue());
                    if (DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        cellText = formatter.format(date);
                    }
                    return cellText;
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case BLANK:
                    return Constant.EMPTY_STRING;
            }

        } catch (Exception e) {
            e.printStackTrace();
            return "row " + rowNum + " or column " + columnName + " does not exist in xls";
        }

        return null;
    }

    public String getCellData(String sheetName, int rowNum, int columnNumber) {
        try {
            if (rowNum < 0)
                return "";

            int index = workbook.getSheetIndex(sheetName);

            if (index == -1)
                return "";

            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(rowNum);
            if (row == null)
                return "";
            cell = row.getCell(columnNumber);
            if (cell == null)
                return "";

            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();

                case FORMULA:
                case NUMERIC:
                    String cellText = String.valueOf(cell.getNumericCellValue());
                    if (DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        cellText = formatter.format(date);
                    }
                    return cellText;

                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());

                case BLANK:
                    return Constant.EMPTY_STRING;

            }
        } catch (Exception e) {

            e.printStackTrace();
            return "row " + rowNum + " or column " + columnNumber + " does not exist  in xls";
        }
        return null;
    }

    public boolean setCellData(String sheetName, int rowNum, String columnName, String data) {
        try {
            fileInputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(fileInputStream);

            if (rowNum <= 0)
                return false;

            int index = workbook.getSheetIndex(sheetName);
            int columnNum = -1;
            if (index == -1)
                return false;

            sheet = workbook.getSheetAt(index);

            row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i).getStringCellValue().trim().equals(columnName))
                    columnNum = i;
            }
            if (columnNum == -1)
                return false;

            sheet.autoSizeColumn(columnNum);
            row = sheet.getRow(rowNum);
            if (row == null)
                row = sheet.createRow(rowNum);

            cell = row.getCell(columnNum);
            if (cell == null)
                cell = row.createCell(columnNum);

            // cell style
//            CellStyle cs = workbook.createCellStyle();
//            cs.setWrapText(true);
//            cell.setCellStyle(cs);
            cell.setCellValue(data);

            fileOutputStream = new FileOutputStream(path);

            workbook.write(fileOutputStream);

            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public boolean addSheet(String sheetname) {

        FileOutputStream fileOut;
        try {
            workbook.createSheet(sheetname);
            fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public boolean removeSheet(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1)
            return false;

        FileOutputStream fileOut;
        try {
            workbook.removeSheetAt(index);
            fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public boolean addColumn(String sheetName, String columnName) {
        try {
            fileInputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(fileInputStream);
            int index = workbook.getSheetIndex(sheetName);
            if (index == -1)
                return false;

            XSSFCellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex());

            sheet = workbook.getSheetAt(index);

            row = sheet.getRow(0);
            if (row == null)
                row = sheet.createRow(0);

            if (row.getLastCellNum() == -1)
                cell = row.createCell(0);
            else
                cell = row.createCell(row.getLastCellNum());

            cell.setCellValue(columnName);
            cell.setCellStyle(style);

            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

        return true;

    }

    public boolean removeColumn(String sheetName, int columnNum) {
        try {
            if (!isSheetExist(sheetName))
                return false;
            fileInputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheet(sheetName);
            XSSFCellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex());
            XSSFCreationHelper createHelper = workbook.getCreationHelper();


            for (int i = 0; i < getRowCount(sheetName); i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    cell = row.getCell(columnNum);
                    if (cell != null) {
                        cell.setCellStyle(style);
                        row.removeCell(cell);
                    }
                }
            }
            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;

    }
}