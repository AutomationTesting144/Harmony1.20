package com.example.a310287808.harmony120;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.Random;
import java.util.concurrent.TimeUnit;
import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/7/2017.
 */

public class ActivityCreation {

    public String IPAddress = "192.168.86.23/api";
    public String HueUserName = "SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2";
    public String HueBridgeParameterType = "groups/0";
    public String finalURL;
    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;
    public String rename;
    MobileElement listItem;
    public int lightCounter=0;
    Dimension size;

    public void ActivityCreation(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {

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
        Object ob = jsonObject.get("action");
        String newString = ob.toString();
        JSONObject jsonObject1 = new JSONObject(newString);
        Object ob1 = jsonObject1.get("on");
        System.out.println(ob1);

        //If the lights in the group are already ON then turn them off
        if (ob1.toString()=="true")
        {
            URL url1 = new URL("http://192.168.86.23/api/SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2/groups/0/action");
            String content = "{"+"\"on\""+":"+"false"+"}";
            HttpURLConnection httpCon = (HttpURLConnection) url1.openConnection();
            httpCon.setDoOutput(true);
            httpCon.setRequestMethod("PUT");
            OutputStreamWriter out = new OutputStreamWriter(httpCon.getOutputStream());
            out.write(content);
            out.close();
            httpCon.getInputStream();
            System.out.println(httpCon.getResponseCode());
        }

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on Activities
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_activities")).click();
        TimeUnit.SECONDS.sleep(2);

        //click on Edit activities
        driver.findElement(By.id("com.logitech.harmonyhub:id/add_edit_textView")).click();
        TimeUnit.SECONDS.sleep(2);

        //click on Add activity
        driver.findElement(By.id("com.logitech.harmonyhub:id/setupEdit")).click();
        TimeUnit.SECONDS.sleep(60);

        //Swiping the screen
        size = driver.manage().window().getSize();
        //Find swipe start and end point from screen's with and height.
        //Find starty point which is at bottom side of screen.
        int starty = (int) (size.height * 0.80);
        //Find endy point which is at top side of screen.
        int endy = (int) (size.height * 0.20);
        //Find horizontal point where you wants to swipe. It is in middle of screen width.
        int startx = size.width / 2;

        //Swipe from Bottom to Top.
        driver.swipe(startx, starty, startx, endy, 3000);
        Thread.sleep(2000);

        //Clicking on lamp icon
        driver.findElement(By.xpath("//android.view.View[@index='40']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on text bar
        driver.findElement(By.id("activityName")).click();
        TimeUnit.SECONDS.sleep(5);

        Random rand = new Random();
        int  n = rand.nextInt() + 1;
        rename = String.valueOf(n);
        driver.findElement(By.id("activityName")).sendKeys(rename);
        TimeUnit.SECONDS.sleep(5);

        //Clicking on forward arrow
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing lights
        driver.findElement(By.xpath("//android.view.View[@bounds='[42,815][147,917]']")).click();
        TimeUnit.SECONDS.sleep(2);
        driver.findElement(By.xpath("//android.view.View[@bounds='[42,1015][147,1116]']")).click();

        //Clicking on forward arrow
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on forward arrow
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing light for activity start
        driver.findElement(By.xpath("//android.view.View[@bounds='[91,700][210,819]']")).click();
        TimeUnit.SECONDS.sleep(5);

        //switching On the light
        driver.findElement(By.id("com.logitech.harmonyhub:id/dd_powerView")).click();
        TimeUnit.SECONDS.sleep(2);

        //forward button
        driver.findElement(By.id("com.logitech.harmonyhub:id/DDC_icon_Device")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing light for activity start
        driver.findElement(By.xpath("//android.view.View[@bounds='[91,1018][210,1137]']")).click();
        TimeUnit.SECONDS.sleep(5);

        //switching On the light
        driver.findElement(By.id("com.logitech.harmonyhub:id/dd_powerView")).click();
        TimeUnit.SECONDS.sleep(2);

        //forward button
        driver.findElement(By.id("com.logitech.harmonyhub:id/DDC_icon_Device")).click();
        TimeUnit.SECONDS.sleep(2);

        //click on forward button
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing activity to END
        driver.findElement(By.xpath("//android.view.View[@content-desc='TestingLamp']")).click();
        TimeUnit.SECONDS.sleep(2);

        //switching On the light
        driver.findElement(By.id("com.logitech.harmonyhub:id/dd_powerView")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking forward button
        driver.findElement(By.id("com.logitech.harmonyhub:id/DDC_icon_Device")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing 2nd light
        driver.findElement(By.xpath("//android.view.View[@content-desc='Light 4']")).click();
        TimeUnit.SECONDS.sleep(2);

        //switching On the light
        driver.findElement(By.id("com.logitech.harmonyhub:id/dd_powerView")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking forward button
        driver.findElement(By.id("com.logitech.harmonyhub:id/DDC_icon_Device")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking button
        driver.findElement(By.id("next")).click();
        TimeUnit.SECONDS.sleep(60);

        //Locate all drop down list elements
        List dropList = driver.findElements(By.id("com.logitech.harmonyhub:id/ActivityName"));
        StringBuffer sb = new StringBuffer();
        for(int i=0; i< dropList.size(); i++){
            listItem = (MobileElement) dropList.get(i);
            System.out.println(listItem.getText());
            Boolean result=listItem.getText().equals(rename);
            if (result==true){
                lightCounter++;
                sb.append(rename);
                sb.append("\n");
            }
            else{
                continue;
            }
        }
        if (lightCounter==0)
        {
            Status = "0";
            ActualResult ="Activity: "+rename+" is not Created";
            Comments = "FAIL: Activity is not created";
            ExpectedResult="Activity: "+rename+" should be created";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {
            Status = "1";
            ActualResult ="Activity: "+rename+" is Created";
            Comments = "NA";
            ExpectedResult="Activity: "+rename+" should be created";
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
        r2c2.setCellValue("17");

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
