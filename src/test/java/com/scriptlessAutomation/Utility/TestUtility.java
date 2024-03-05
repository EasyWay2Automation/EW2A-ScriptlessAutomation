package com.scriptlessAutomation.Utility;


import com.scriptlessAutomation.Driver.ExecuteTest;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;


public class TestUtility extends ExecuteTest {

    public static void deleteFileORFolder(String fileAndFolderPath) {

        File index = new File(fileAndFolderPath);
        if(index.exists()) {
            index.delete();
        }

    }

    public static void waitForElement(WebElement element) {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(Constant.EXPLICITTIMEOUTLIMIT));
        wait.until(ExpectedConditions.visibilityOf(element));
    }

    public static void scrollToElement(WebElement element) {
        waitForElement(element);
        ((JavascriptExecutor)driver).executeScript("arguments[0].scrollIntoView(true);", element);

    }

    public static String getCurrentDateAndTime() {

        return new SimpleDateFormat("dd_MM_yyyy_hh_mm_ss").format(new Date());
    }

    public static void createFolders(String folderPath) {

        File index = new File(folderPath);
        if(!index.exists()) {
            index.mkdirs();
        }

    }
}
