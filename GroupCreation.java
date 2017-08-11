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
import java.util.List;
import java.util.Random;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/3/2017.
 */

public class GroupCreation {
    public String ActualResult;
    public String ExpectedResult;
    public String Status;
    public String Comments;
    public String RoomName;
    MobileElement listItem;
    public int lightCounter=0;

    public void GroupCreation (AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException  {

        driver.navigate().back();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on device
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_devices")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking on Edit Devices button
        driver.findElement(By.id("com.logitech.harmonyhub:id/add_edit_textView")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on +group button
        driver.findElement(By.id("com.logitech.harmonyhub:id/addGroupLayout")).click();
        TimeUnit.SECONDS.sleep(15);

        //Entering new name
        Random rand = new Random();
        int  n = rand.nextInt() + 1;
        RoomName = String.valueOf(n);
        driver.findElement(By.id("groupNameText")).sendKeys(RoomName);
        TimeUnit.SECONDS.sleep(5);
        driver.hideKeyboard();

        //Choosing Lights
        driver.findElement(By.xpath("//android.view.View[@bounds='[42,885][147,987]']")).click();
        TimeUnit.SECONDS.sleep(5);
        driver.findElement(By.xpath("//android.view.View[@bounds='[42,1085][147,1186]']")).click();
        TimeUnit.SECONDS.sleep(5);

        //Clicking on forward arrow
        driver.findElement(By.id("done")).click();
        TimeUnit.SECONDS.sleep(10);

        //Choosing the group to turn it on
        driver.findElement(By.id("com.logitech.harmonyhub:id/device_right_icon")).click();
        TimeUnit.SECONDS.sleep(5);

        //Locate all drop down list elements
        List dropList = driver.findElements(By.id("com.logitech.harmonyhub:id/device_name"));
        StringBuffer sb = new StringBuffer();
        for(int i=0; i< dropList.size(); i++){
            listItem = (MobileElement) dropList.get(i);
            Boolean result=listItem.getText().equals(RoomName);
            if (result==true){
                lightCounter++;
                sb.append(RoomName);
                sb.append("\n");
            }
            else{
                continue;
            }
        }
        if (lightCounter==0)
        {
            Status = "0";
            ActualResult ="Group: "+RoomName+" is not Created";
            Comments = "FAIL: Group is not created";
            ExpectedResult="Group: "+RoomName+" should be created";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {
            Status = "1";
            ActualResult ="Group: "+RoomName+" is Created";
            Comments = "NA";
            ExpectedResult="Group: "+RoomName+" should be created";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }

        //Going back on home from Harmony
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
        r2c2.setCellValue("12");

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
