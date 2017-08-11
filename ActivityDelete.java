package com.example.a310287808.harmony120;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.openqa.selenium.By;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/8/2017.
 */

public class ActivityDelete {
    public String ActualResult;
    public String ExpectedResult;
    public String Status;
    public String Comments;
    MobileElement listItem;
    public String lightValue;
    public int lightCounter=0;

    public void ActivityDelete (AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException  {

        driver.navigate().back();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on Menu button
        driver.findElement(By.id("com.logitech.harmonyhub:id/left_command")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on harmony setup
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony Setup']")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking on Add/Edit devices
        driver.findElement(By.xpath("//android.widget.TextView[@text='Add/Edit Devices & Activities']")).click();
        TimeUnit.SECONDS.sleep(30);

        //click on Groups
        driver.findElement(By.xpath("//android.view.View[@content-desc='ACTIVITIES']")).click();
        TimeUnit.SECONDS.sleep(5);

        //choosing activity to delete
        driver.findElement(By.xpath("//android.view.View[@bounds='[276,420][1200,500]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //click on edit button to get the name of activity
        driver.findElement(By.id("editIcon")).click();
        TimeUnit.SECONDS.sleep(2);

        lightValue=driver.findElement(By.id("activityName")).getText();
        System.out.println(lightValue);

        //clicking on back icon
        driver.findElement(By.id("back")).click();
        TimeUnit.SECONDS.sleep(2);

        //Delete the activity
        driver.findElement(By.id("btnDelActivity")).click();
        TimeUnit.SECONDS.sleep(2);

        //Delete the activity
        driver.findElement(By.id("btnDelete")).click();
        Boolean result=driver.findElement(By.xpath("//android.view.View[@content-desc='Syncing Your Changes']")).isDisplayed();
        System.out.println("waiting" +result);
        TimeUnit.SECONDS.sleep(60);

        //forward button
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(2);

        //setting button
        driver.findElement(By.id("close")).click();
        TimeUnit.SECONDS.sleep(2);

        //System.out.println(lightCounter);
        if (result==true)
        {

            Status = "1";
            ActualResult ="Activity: " + lightValue + " is Deleted";
            Comments = "NA";
            ExpectedResult="Activity: "+lightValue+" should be Deleted";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {

            Status = "0";
            ActualResult ="Activity: " + lightValue + " is not Deleted";
            Comments = "FAIL: Activity is not deleted";
            ExpectedResult="Activity: "+lightValue+" should be Deleted";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }

        //Going back on home from Harmony
        driver.navigate().back();
        driver.navigate().back();
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);
    }

    public String CurrentdateTime;
    public int nextRowNumber;
    public void storeResultsExcel(String excelStatus, String excelActualResult, String excelComments, String resultFileName, String ExcelExpectedResult
            ,String resultAPIVersion, String resultSWVersion) throws IOException {

        Calendar cal = Calendar.getInstance();
        SimpleDateFormat sdf  = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        CurrentdateTime = sdf.format(cal.getTime());
        FileInputStream fsIP = new FileInputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
        nextRowNumber=workbook.getSheetAt(0).getLastRowNum();
        nextRowNumber++;
        HSSFSheet sheet = workbook.getSheetAt(0);

        HSSFRow row2 = sheet.createRow(nextRowNumber);
        HSSFCell r2c1 = row2.createCell((short) 0);
        r2c1.setCellValue(CurrentdateTime);

        HSSFCell r2c2 = row2.createCell((short) 1);
        r2c2.setCellValue("21");

        HSSFCell r2c3 = row2.createCell((short) 2);
        r2c3.setCellValue(excelStatus);

        HSSFCell r2c4 = row2.createCell((short) 3);
        r2c4.setCellValue(excelActualResult);

        HSSFCell r2c5 = row2.createCell((short) 4);
        r2c5.setCellValue(excelComments);

        HSSFCell r2c6 = row2.createCell((short) 5);
        r2c6.setCellValue(resultAPIVersion);

        HSSFCell r2c7 = row2.createCell((short) 6);
        r2c7.setCellValue(resultSWVersion);

        fsIP.close();
        FileOutputStream out =
                new FileOutputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        workbook.write(out);
        out.close();

    }
}
