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
import java.util.concurrent.TimeUnit;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/7/2017.
 */

public class GroupDeletion {

    public String ActualResult;
    public String ExpectedResult;
    public String Status;
    public String Comments;
    MobileElement listItem;
    public String rename;
    public int lightCounter=0;

    public void GroupDeletion (AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException  {

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
        driver.findElement(By.xpath("//android.view.View[@content-desc='GROUPS']")).click();
        TimeUnit.SECONDS.sleep(5);

        //click on lights
        driver.findElement(By.xpath("//android.view.View[@bounds='[0,304][1440,595]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking lights to rename
        driver.findElement(By.xpath("//android.view.View[@bounds='[0,665][1440,934]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //editing the group name
        driver.findElement(By.id("editGroupButton")).click();
        TimeUnit.SECONDS.sleep(2);

        //getting the group name
        driver.findElement(By.id("groupNameText")).click();
        String lightValue=driver.findElement(By.id("groupNameText")).getText();

//        //Clearing the old name
//        driver.findElement(By.id("groupNameText")).clear();
//        TimeUnit.SECONDS.sleep(2);
//        //Entering new name
//        Random rand = new Random();
//        int  n = rand.nextInt() + 1;
//        rename = String.valueOf(n);
//        driver.findElement(By.id("groupNameText")).sendKeys(rename);
//        TimeUnit.SECONDS.sleep(5);
//        driver.hideKeyboard();

        //clicking on forward arrow
        driver.findElement(By.id("done")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on DELETE button
        driver.findElement(By.id("deleteGroupButton")).click();
        TimeUnit.SECONDS.sleep(2);

        //Confirming the deletion
        driver.findElement(By.xpath("//android.widget.Button[@content-desc='YES']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on back arrow
        driver.findElement(By.id("back")).click();
        TimeUnit.SECONDS.sleep(2);

        driver.findElement(By.id("close")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking on Device
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_devices")).click();
        TimeUnit.SECONDS.sleep(2);

        //spinner arrow
        driver.findElement(By.id("com.logitech.harmonyhub:id/device_right_icon")).click();
        TimeUnit.SECONDS.sleep(2);

        List dropList = driver.findElements(By.id("com.logitech.harmonyhub:id/device_name"));
        StringBuffer sb = new StringBuffer();
        for(int i=0; i< dropList.size(); i++){
            listItem = (MobileElement) dropList.get(i);
            //System.out.println(listItem.getText());
            Boolean result=listItem.getText().equals(lightValue);
            // System.out.println(result);
            if (result==true){
                lightCounter++;
                sb.append(lightValue);
                sb.append("\n");
            }
            else{
                continue;
            }
        }
        //System.out.println(lightCounter);
        if (lightCounter==0)
        {

            Status = "1";
            ActualResult ="Group: " + lightValue + " is Deleted";
            Comments = "NA";
            ExpectedResult="Group: "+lightValue+" should be Deleted";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {

            Status = "0";
            ActualResult ="Group: " + lightValue + " is not Deleted";
            Comments = "FAIL: Group is not deleted";
            ExpectedResult="Group: "+lightValue+" should be Deleted";
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
        r2c2.setCellValue("14");

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
