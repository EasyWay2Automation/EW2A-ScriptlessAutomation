# EW2AScriptlessAutomation
EW2A stands for "EasyWay2Automation" and is a keyword driven and script less automation framework designed using Java programming language.

Major Advantages of EW2A Script less Automation Framework :
1. Involvement of Functional Testers
2. To avoid writing Test Cases
3. Less Code Maintenance
4. Less Maintenance of objects
5. Reduce effort from Test Design and Test Execution Phase
6. Switching between Application Under Tests
7. No Use Of Third Party Reporting Tool
8. No Use of Built-in or Third Party Testing Framework

When ever a new test case needs to be developed then the tester has to make an entry in the Excel sheet.

Controller Sheet :
Controller sheet controls the test cases to be executed. If the cell value of ‘Run Mode’ column is ‘Y’ then the corresponding test case will be considered for execution.

Test Case Sheets:
NOTE : Excel workbook may have multiple test case sheets.
Operation : This column contains a list of keywords (Actions to be on UI). For each keyword, corresponding code has been written in “UIOperation.java”
Object Name : This is a unique name to identify the object name in the “objects.properties” file
Object Type: A list of locators used in Selenium to identify objects uniquely on a web page.
Value: This is optional. If you want a particular value to be selected from a drop down or to be entered in an edit field or to verify text on a web page then you must enter a value.
