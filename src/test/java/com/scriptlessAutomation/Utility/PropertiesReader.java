package com.scriptlessAutomation.Utility;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

public class PropertiesReader {

    Properties p = new Properties();

    /** This function reads the objects.properties file and returns the reference variable of Properties class
     * Parameters : NA
     * Return Type : Properties */

    public Properties getLocators() {
        try {
            FileReader reader = new FileReader(System.getProperty("user.dir") + "\\src\\test\\resources\\config\\objects.properties");
            try {
                p.load(reader);
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return p;
    }

}
