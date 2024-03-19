package com.scriptlessAutomation.Driver;

import com.scriptlessAutomation.Utility.Constant;
import com.scriptlessAutomation.Utility.ExcelUtility;
import com.scriptlessAutomation.Utility.PropertiesReader;
import com.scriptlessAutomation.Utility.TestUtility;
import com.scriptlessAutomation.operations.UIOperation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;

import java.net.InetAddress;
import java.util.Properties;

public class ExecuteTest {

    static ExcelUtility excel = null;
    static PropertiesReader reader = null;
    static UIOperation operation = null;
    public static String testStepStatus = null;
    public static WebDriver driver = null;
    public static String currentScreenshotPath = null;
    static Properties allObjects = null;
    static int testCaseCount = 0;

    public static void main(String[] args) {

        try {
            System.out.println("#################################################");
            System.out.println("SCRIPT LESS AUTOMATION FRAMEWORK (EW2A) : STARTED");
            System.out.println("#################################################");
            System.out.println();
            System.out.println("********************************************");
            System.out.println("Project Name : ScriptlessAutomation");
            System.out.println("Java Version : " + System.getProperty("java.version"));
            System.out.println("OS : " + System.getProperty("os.name"));
            System.out.println("User : " + System.getProperty("user.name"));

            InetAddress myHost = InetAddress.getLocalHost();
            System.out.println("Hostname : " + myHost.getHostName());

            System.out.println("********************************************");

            /* Initialize the Reference variable Of ExcelUtility Class */
            excel = new ExcelUtility(Constant.filepath, Constant.filename);
            /* Initialize the Reference variable Of UIOperation Class */
            operation = new UIOperation();
            /* Initialize the Reference variable Of PropertiesReader Class */
            reader = new PropertiesReader();
            allObjects = reader.getLocators();

            Sheet controllerSheet = excel.getSheet("Controller");

            /* Count Number Of Test Cases available in the Controller Sheet */
            testCaseCount = excel.getRowCount("Controller");
            System.out.println("Test Case Count : " + (testCaseCount-4));
            int tcColumnCount = excel.getColumnCount("Controller");

            /* Create a column as 'Status' in the Controller Sheet and customize it */
            Row controllerRow = controllerSheet.getRow(4);
            Cell cell = controllerRow.createCell(tcColumnCount);
            cell.setCellValue("Status");
            CellStyle style = excel.getCellStyle("Header");
            cell.setCellStyle(style);

            /* Find out the test cases for which Run Mode is 'Y' */
            for (int i = 5; i <= testCaseCount; i++) {

                controllerRow = controllerSheet.getRow(i);
                cell = controllerRow.getCell(tcColumnCount - 1);
                String runMode = cell.getStringCellValue().trim();
                if (runMode.equalsIgnoreCase("Y")) {
                    cell = controllerRow.getCell(0);
                    String testCaseID = cell.getStringCellValue().trim();
                    System.out.println("=============================================");
                    System.out.println("Test Case ID : " + testCaseID);
                    System.out.println("=============================================");

                    /* Create a Folder Named By Test Case ID to save screenshots */
                    currentScreenshotPath = Constant.screenshotPath + "\\" + testCaseID + "_" + TestUtility.getCurrentDateAndTime();
                    TestUtility.createFolders(currentScreenshotPath);

                    Sheet testCaseSheet = excel.getSheet(testCaseID);
                    int tsColumnCount = excel.getColumnCount(testCaseID);
                    Row testCaseRow = testCaseSheet.getRow(0);
                    cell = testCaseRow.createCell(tsColumnCount);
                    cell.setCellValue("Status");
                    style = excel.getCellStyle("Header");
                    cell.setCellStyle(style);

                    cell = testCaseRow.createCell(tsColumnCount + 1);
                    cell.setCellValue("Screenshots");
                    style = excel.getCellStyle("Header");
                    cell.setCellStyle(style);

                    String testCaseStatus = "Passed";

                    int testStepsCount = excel.getRowCount(testCaseID);
                    for (int j = 1; j <= testStepsCount; j++) {
                        testCaseRow = testCaseSheet.getRow(j);
                        testStepStatus = operation.perform(allObjects, testCaseRow.getCell(0).getStringCellValue().trim(),
                                testCaseRow.getCell(1).getStringCellValue().trim(), testCaseRow.getCell(2).getStringCellValue().trim(),
                                testCaseRow.getCell(3).getStringCellValue().trim());

                        System.out.println(testCaseRow.getCell(0).getStringCellValue().trim() + "---"
                                + testCaseRow.getCell(1).getStringCellValue().trim() + "---"
                                + testCaseRow.getCell(2).getStringCellValue().trim() + "---"
                                + testCaseRow.getCell(3).getStringCellValue().trim() + "---" + testStepStatus);


                        // Update Test Step Status
                        cell = testCaseRow.createCell(tsColumnCount);
                        cell.setCellValue(testStepStatus);
                        style = excel.getCellStyle(testStepStatus);
                        cell.setCellStyle(style);

                        if (testCaseRow.getCell(0).getStringCellValue().trim().equals("Take Screenshot")) {
                            cell = testCaseRow.createCell(tsColumnCount + 1);
                            cell.setCellValue(currentScreenshotPath);
                        }

                        if (testStepStatus.equals("Failed")) {
                            testCaseStatus = "Failed";
                            for (int k = j + 1; k <= testStepsCount; k++) {

                                // Update Test Step Status as 'Skipped'
                                testCaseRow = testCaseSheet.getRow(k);
                                cell = testCaseRow.createCell(tsColumnCount);
                                cell.setCellValue("Skipped");
                                style = excel.getCellStyle("Skipped");
                                cell.setCellStyle(style);

                            }
                            break;
                        }

                    }

                    // Update Test Case Status

                    cell = controllerSheet.getRow(i).createCell(tcColumnCount);
                    cell.setCellValue(testCaseStatus);
                    style = excel.getCellStyle(testCaseStatus);
                    cell.setCellStyle(style);


                }

            }


        } catch (Exception ex) {

            System.out.println("Exception occurred : " + ex.getMessage());
            ex.printStackTrace();
            testStepStatus = operation.perform(allObjects, "Close Browser", "", "", "");
            System.out.println("Close Browser --- " + testStepStatus);

        } finally {
	    for(int i=0; i<=testCaseCount-4; i++){
                excel.autoFitColumn(i);
            }
            excel.generateReport(Constant.excelReportPath);
            excel.closeWorkbook();

            System.out.println("##################################################");
            System.out.println("SCRIPT LESS AUTOMATION FRAMEWORK (EW2A) : COMPLETE");
            System.out.println("##################################################");
        }
    }
}
