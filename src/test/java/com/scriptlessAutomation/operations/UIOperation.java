package com.scriptlessAutomation.operations;

import com.google.common.io.Files;
import com.scriptlessAutomation.Driver.ExecuteTest;
import com.scriptlessAutomation.Utility.Constant;
import com.scriptlessAutomation.Utility.TestUtility;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.time.Duration;
import java.util.Properties;
import java.util.Set;

public class UIOperation extends ExecuteTest {
    Actions builder;
    String popupWindowHandle = null;
    String parentWindowHandle = null;
    Set<String> handles;

    public String perform(Properties p,String operation,String objectName,String objectType,String value ) {

        testStepStatus = "Passed";

        try {

            switch(operation) {

                case "Open Browser":
                    if(objectName.equalsIgnoreCase("CHROME")) {
                        driver = new ChromeDriver();
                        driver.manage().window().maximize();
                    }

                    driver.manage().window().maximize();
                    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
                    driver.get(value);

                    break;

                case "Click Element":
                    WebElement elementToBeClicked = driver.findElement(this.getObject(p,objectType,objectName));
                    TestUtility.scrollToElement(elementToBeClicked);
                    elementToBeClicked.click();
                    break;
                case "Input Text":
                    WebElement elementToBeSet = driver.findElement(this.getObject(p, objectType, objectName));
                    TestUtility.scrollToElement(elementToBeSet);
                    elementToBeSet.sendKeys(value);
                    break;

                case "Select By Visible Text":
                    WebElement elementToBeSelected = driver.findElement(this.getObject(p,objectType,objectName));
                    TestUtility.scrollToElement(elementToBeSelected);
                    Select select = new Select(elementToBeSelected);
                    highlightElement(elementToBeSelected);
                    select.selectByVisibleText(value);
                    break;

                case "Select By Visible Value":
                    WebElement elementToBeSelectedByValue = driver.findElement(this.getObject(p,objectType,objectName));
                    TestUtility.scrollToElement(elementToBeSelectedByValue);
                    highlightElement(elementToBeSelectedByValue);
                    Select selectByval = new Select(elementToBeSelectedByValue);
                    selectByval.selectByValue(value);
                    break;

                case "Select By Index":
                    WebElement elementToBeSelectedByIndex = driver.findElement(this.getObject(p,objectType,objectName));
                    TestUtility.scrollToElement(elementToBeSelectedByIndex);
                    highlightElement(elementToBeSelectedByIndex);
                    Select selectByInd = new Select(elementToBeSelectedByIndex);
                    selectByInd.selectByIndex(Integer.parseInt(value));
                    break;

                case "Verify Element Present":
                    WebElement elementToBeVerified = driver.findElement(this.getObject(p,objectType,objectName));
                    if(!elementToBeVerified.isDisplayed()) {
                        testStepStatus = "Failed";
                    }else {
                        TestUtility.scrollToElement(elementToBeVerified);
                        highlightElement(elementToBeVerified);
                    }
                    break;

                case "Verify Element Enabled":
                    WebElement elementToBeChecked = driver.findElement(this.getObject(p,objectType,objectName));
                    if(!elementToBeChecked.isEnabled()) {
                        testStepStatus = "Failed";
                    }else {
                        TestUtility.scrollToElement(elementToBeChecked);
                        highlightElement(elementToBeChecked);
                    }
                    break;

                case "Java Script Click":
                    WebElement elementToBeClickedUsingJS = driver.findElement(this.getObject(p,objectType,objectName));
                    ((JavascriptExecutor)driver).executeScript("arguments[0].click();",elementToBeClickedUsingJS);
                    break;

                case "Static Wait":
                    long waitTime = Long.parseLong(value);
                    Thread.sleep(waitTime);
                    break;

                case "Navigate To URL":
                    driver.navigate().to(value);
                    break;

                case "Verify Text":
                    WebElement element = driver.findElement(this.getObject(p, objectType, objectName));
                    highlightElement(element);
                    String actText = element.getText().trim();
                    if(!actText.equals(value)){
                        testStepStatus = "Failed";
                    }
                    break;

                case "Take Screenshot":
                    TakesScreenshot srcShot = (TakesScreenshot)driver;
                    File srcFile = srcShot.getScreenshotAs(OutputType.FILE);
                    File destFile = new File(currentScreenshotPath+"\\"+TestUtility.getCurrentDateAndTime()+".png");
                    Files.copy(srcFile, destFile);
                    break;

                case "Mouse Over":
                    builder = new Actions(driver);
                    WebElement elementOnMouseOver = driver.findElement(this.getObject(p, objectType, objectName));
                    TestUtility.waitForElement(elementOnMouseOver);
                    builder.moveToElement(elementOnMouseOver).perform();
                    break;

                case "Right Click":
                    builder = new Actions(driver);
                    WebElement elementToBeRightClicked = driver.findElement(this.getObject(p, objectType, objectName));
                    TestUtility.waitForElement(elementToBeRightClicked);
                    builder.contextClick(elementToBeRightClicked).perform();
                    break;

                case "Double Click":
                    builder = new Actions(driver);
                    WebElement elementToBeDoubleClicked = driver.findElement(this.getObject(p, objectType, objectName));
                    TestUtility.waitForElement(elementToBeDoubleClicked);
                    builder.doubleClick(elementToBeDoubleClicked).perform();
                    break;

                case "Switch To Child Window":
                    parentWindowHandle = driver.getWindowHandle();
                    System.out.println("Parent Window : "+parentWindowHandle);
                    handles = driver.getWindowHandles();
                    System.out.println(handles);
                    for(String handle : handles) {
                        if(!handle.equals(parentWindowHandle)) {
                            popupWindowHandle = handle;
                        }
                    }

                    driver.switchTo().window(popupWindowHandle);
                    System.out.println("Driver has switched to Child Window...");
                    break;

                case "Handle Alert":
                    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(Constant.EXPLICITTIMEOUTLIMIT));
                    Alert alert = wait.until(ExpectedConditions.alertIsPresent());
                    System.out.println("Alert Text : "+alert.getText());
                    alert.accept();
                    break;

                case "Switch To Parent Window":
                    driver.switchTo().window(parentWindowHandle);
                    break;

                case "Switch To Frame":
                    driver.switchTo().frame(driver.findElement(this.getObject(p, objectType, objectName)));
                    break;

                case "Switch To Parent Page":
                    driver.switchTo().defaultContent();
                    break;
                case "Close Browser":
                    if(driver!=null) {
                        driver.quit();
                    }
                    break;

            }

        }catch(Exception ex) {
            testStepStatus = "Failed";
            System.out.println("Exception Occurred while performing UI operation " + ex.getMessage());
            if(driver!=null) {
                driver.quit();
            }
        }

        return testStepStatus;
    }


    private By getObject(Properties p, String objectType, String objectName) {
        switch (objectType) {
            case "ID":
                return By.id(p.getProperty(objectName));
            case "XPATH":
                return By.xpath(p.getProperty(objectName));
            case "NAME":
                return By.name(p.getProperty(objectName));
            case "CSS SELECTOR":
                return By.cssSelector(p.getProperty(objectName));
            case "CLASS NAME":
                return By.className(p.getProperty(objectName));
            case "LINK TEXT":
                return By.linkText(p.getProperty(objectName));
            case "PARTIAL LINK TEXT":
                return By.partialLinkText(p.getProperty(objectName));
            case "TAG NAME":
                return By.tagName(p.getProperty(objectName));
            default:
                System.out.println("Wrong Object Type");
                return null;
        }
    }

    private void highlightElement(WebElement element) {
        JavascriptExecutor js = (JavascriptExecutor)driver;
        js.executeScript("arguments[0].setAttribute('style','background: yellow; border: 2px solid red;')",element);
    }
}
