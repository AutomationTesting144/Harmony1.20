package com.example.a310287808.harmony120;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;
import org.openqa.selenium.By;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/8/2017.
 */

public class ActivityChange {
    public String IPAddress = "192.168.86.23/api";
    public String HueUserName = "SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2";
    public String HueBridgeParameterType = "lights/37";
    public String finalURL;
    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;
    public String rename;
    public String LightName;

    public void ActivityChange(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {

        driver.navigate().back();
        driver.navigate().back();
        HttpURLConnection connection;
        finalURL = "http://" + IPAddress + "/" + HueUserName + "/" + HueBridgeParameterType;
        URL url = new URL(finalURL);
        connection = (HttpURLConnection) url.openConnection();
        connection.connect();
        InputStream stream = connection.getInputStream();
        BufferedReader reader = new BufferedReader(new InputStreamReader(stream));
        StringBuffer br = new StringBuffer();
        String line = " ";
        while ((line = reader.readLine()) != null) {
            br.append(line);
        }

        String output1 = br.toString();
        JSONObject jsonObject = new JSONObject(output1);
        Object ob = jsonObject.get("state");
        String newString = ob.toString();
        JSONObject jsonObject1 = new JSONObject(newString);
        Object ob1 = jsonObject1.get("on");

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

        //click on Activities
        driver.findElement(By.xpath("//android.view.View[@content-desc='ACTIVITIES']")).click();
        TimeUnit.SECONDS.sleep(5);

        //clicking on activities to edit
        driver.findElement(By.xpath("//android.view.View[@bounds='[1312,420][1410,518]']")).click();
        TimeUnit.SECONDS.sleep(5);

        //clicking on Edit start sequence
        driver.findElement(By.xpath("//android.view.View[@content-desc='EDIT START SEQUENCE']")).click();
        TimeUnit.SECONDS.sleep(5);

        //Edit home control
        driver.findElement(By.id("addHCDevices")).click();
        TimeUnit.SECONDS.sleep(2);

        //Selecting light to delete
        driver.findElement(By.xpath("//android.view.View[@content-desc='TestingLamp']")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking on forward arrow
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(5);

        //clicking on forward arrow
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(5);

        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(5);

        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(5);

        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(5);

        //clicking on Settings cross mark
        driver.findElement(By.id("close")).click();
        TimeUnit.SECONDS.sleep(5);

        //clicking on Activities
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_activities")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing activities for switching ON
        driver.findElement(By.xpath("//android.widget.ImageView[@bounds='[81,464][220,608]']")).click();
        TimeUnit.SECONDS.sleep(10);

        //getting the name of the activity
        rename=driver.findElement(By.xpath("//android.widget.TextView[@bounds='[382,478][1440,594]']")).getText();
        TimeUnit.SECONDS.sleep(2);


        driver.navigate().back();

        finalURL = "http://" + IPAddress + "/" + HueUserName + "/" + HueBridgeParameterType;
        URL url3 = new URL(finalURL);
        connection = (HttpURLConnection) url3.openConnection();
        connection.connect();
        InputStream stream3 = connection.getInputStream();
        BufferedReader reader3 = new BufferedReader(new InputStreamReader(stream3));
        StringBuffer br3 = new StringBuffer();
        String line3 = " ";
        while ((line3 = reader3.readLine()) != null) {
            br3.append(line3);
        }

        String output3 = br3.toString();
        JSONObject jsonObject3 = new JSONObject(output3);
        Object ob3 = jsonObject3.get("state");
        String newString3 = ob3.toString();
        JSONObject jsonObject4 = new JSONObject(newString3);
        Object ob4 = jsonObject4.get("on");
        JSONObject jsonObject5 = new JSONObject(output3);
        Object ob5 = jsonObject5.get("name");
        LightName=ob5.toString();


        if (ob4.toString().equals(ob1.toString()))
        {
            Status = "1";
            ActualResult ="Light: "+LightName+" is removed from the activity: "+rename+ " and is not controlled by the activity";
            Comments = "NA";
            ExpectedResult="Light should be removed and should not be controlled by activity";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {
            Status = "0";
            ActualResult ="Light: "+LightName+" is not removed from the activity: "+rename+ " and is controlled by the activity";
            Comments = "FAIL: status of the removed light is: "+ob4.toString();
            ExpectedResult="Light should be removed and should not be controlled by activity";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }
        driver.navigate().back();
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);
    }


    public String CurrentdateTime;
    public int nextRowNumber;
    public void storeResultsExcel (String excelStatus, String excelActualResult, String excelComments, String resultFileName, String ExcelExpectedResult, String resultAPIVersion, String resultSWVersion) throws IOException {
        Calendar cal = Calendar.getInstance();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        CurrentdateTime = sdf.format(cal.getTime());
        FileInputStream fsIP = new FileInputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
        nextRowNumber = workbook.getSheetAt(0).getLastRowNum();
        nextRowNumber++;
        HSSFSheet sheet = workbook.getSheetAt(0);

        HSSFRow row2 = sheet.createRow(nextRowNumber);
        HSSFCell r2c1 = row2.createCell((short) 0);
        r2c1.setCellValue(CurrentdateTime);

        HSSFCell r2c2 = row2.createCell((short) 1);
        r2c2.setCellValue("20");

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
        FileOutputStream out =new FileOutputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        workbook.write(out);
        out.close();

    }
}
