package com.scriptlessAutomation.Utility;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtility {
    Workbook workbook = null;
    Sheet sheet = null;
    Row row = null;
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
        return sheet.getLastRowNum();
    }

    /** Function to return number of columns of a specified sheet
     * Parameter : sheetName
     * Return Type : Integer */

    public int getColumnCount(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(4);
        return row.getLastCellNum();
    }

    /** Function to return number of rows of a specified sheet
     * Parameter : sheetNo
     * Return Type : Integer */

    public int getRowCount(int sheetNo) {
        sheet = workbook.getSheetAt(sheetNo);
        return sheet.getLastRowNum();
    }

    /** Function to return number of columns of a specified sheet
     * Parameter : sheetNo
     * Return Type : Integer */

    public int getColumnCount(int sheetNo) {
        sheet = workbook.getSheetAt(sheetNo);
        row = sheet.getRow(4);
        return row.getLastCellNum();
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
     * Parameter : reportPath
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


    /** Function to Close Workbook
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

    /**
     * Function to set column width of a specified sheet
     * @param sheetNo
     * Return Type: void
     */
    public void autoFitColumn(int sheetNo){
        sheet = workbook.getSheetAt(sheetNo);
        for(int i=0; i<getColumnCount(0) ; i++){
            sheet.autoSizeColumn(i);
        }

    }

    /**
     * Function to insert an image into a specified cell
     * @param sheetName
     * @param screenShotPath
     * @param row1
     * @param col1
     * @param row2
     * @param col2
     * Return Type: void
     */
    public void attachImageInToCell(String sheetName, String screenShotPath, int row1, int col1, int row2, int col2){
        try {
            sheet = workbook.getSheet(sheetName);
            InputStream imgStream  = new FileInputStream(screenShotPath);
            byte[] imgBytes = IOUtils.toByteArray(imgStream);
            int pictureIdx = workbook.addPicture(imgBytes, Workbook.PICTURE_TYPE_PNG);
            imgStream.close();
            Drawing drawing = sheet.createDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            anchor.setCol1(col1);
            anchor.setRow1(row1);
            anchor.setCol2(col2);
            anchor.setRow2(row2);
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize();

        }catch(Exception ex){
            ex.printStackTrace();
        }


    }
}
