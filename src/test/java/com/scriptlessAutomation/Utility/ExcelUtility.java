package com.scriptlessAutomation.Utility;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtility {
    Workbook workbook = null;
    Sheet sheet = null;
    Row row = null;
    Cell cell = null;
    File file = null;
    FileInputStream inputstream = null;

    /** Constructor to initialize reference variable of Workbook Interface
     *  Parameters : filePath, fileName
     *  Return Type : NA */
    public ExcelUtility(String filePath, String fileName) {
        try {
            file = new File(filePath+"\\"+fileName);
            inputstream = new FileInputStream(file);
            String fileExtension = fileName.substring(fileName.indexOf("."));
            try {
                if(fileExtension.equals(".xlsx")) {
                    workbook = new XSSFWorkbook(inputstream);
                }else if(fileExtension.equals(".xls")) {
                    workbook = new HSSFWorkbook(inputstream);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

    /** Function to return reference variable of Sheet Interface
     * Parameter : sheetName
     * Return Type : Sheet */

    public Sheet getSheet(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        return sheet;
    }


    /** Function to return number of rows of a specified sheet
     * Parameter : sheetName
     * Return Type : Integer */

    public int getRowCount(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        int rowCount = sheet.getLastRowNum();
        return rowCount;
    }

    /** Function to return number of columns of a specified sheet
     * Parameter : sheetName
     * Return Type : Integer */

    public int getColumnCount(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(4);
        int colCount = row.getLastCellNum();
        return colCount;
    }

    /** Function to customize Cell based on specified Status
     * Parameter : status
     * Return Type : CellStyle */

    public CellStyle getCellStyle(String status) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        if(status.equalsIgnoreCase("Header")) {
            style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = workbook.createFont();
            font.setBold(true);
            font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
            style.setFont(font);
        }else if(status.equalsIgnoreCase("Passed")) {
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }else if(status.equalsIgnoreCase("Failed")) {
            style.setFillForegroundColor(IndexedColors.RED.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }else if(status.equalsIgnoreCase("Skipped")) {
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        return style;
    }

    /** Function to generate Report
     * Parameter : NA
     * Return Type : void */

    public void generateReport(String reportPath) {
        try {
            file = new File(reportPath);
            FileOutputStream outputstream = new FileOutputStream(file);
            workbook.write(outputstream);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }


    /** Function to generate Report
     * Parameter : NA
     * Return Type : void */

    public void closeWorkbook() {
        try {
            inputstream.close();
            workbook.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
