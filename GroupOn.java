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
import java.io.OutputStreamWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/7/2017.
 */

public class GroupOn {

    public String IPAddress = "192.168.86.23/api";
    public String HueUserName = "SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2";
    public String HueBridgeParameterType = "groups/0";
    public String RoomName;
    public String finalURL;
    public String Status;
    public String ActualResult;
    public String Comments;
    public String ExpectedResult;

    public void GroupOn(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {
        //Getting the state of group from API
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
        Object ob1 = jsonObject1.get("all_on");
        System.out.println(ob1);

        //If the lights in the group are already ON then turn them off
        if (ob1.toString()=="true")
        {
            URL url1 = new URL("http://192.168.86.23/api/SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2/groups/0/state");
            String content = "{"+"\"all_on\""+":"+"false"+"}";
            HttpURLConnection httpCon = (HttpURLConnection) url1.openConnection();
            httpCon.setDoOutput(true);
            httpCon.setRequestMethod("PUT");
            OutputStreamWriter out = new OutputStreamWriter(httpCon.getOutputStream());
            out.write(content);
            out.close();
            httpCon.getInputStream();
            System.out.println(httpCon.getResponseCode());

        }
        //Clicking on Home button of the device
        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on device
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_devices")).click();
        TimeUnit.SECONDS.sleep(2);

        //spinner arrow
        driver.findElement(By.id("com.logitech.harmonyhub:id/device_right_icon")).click();
        TimeUnit.SECONDS.sleep(2);

        //get the name of the group
        RoomName= driver.findElement(By.xpath("//android.widget.TextView[@bounds='[301,773][1265,861]']")).getText();

        //Clicking on Group to turn on
        driver.findElement(By.xpath("//android.widget.ImageView[@bounds='[78,773][222,917]']")).click();
        TimeUnit.SECONDS.sleep(10);
        driver.navigate().back();

        //Checking the status of lights

        finalURL = "http://" + IPAddress + "/" + HueUserName + "/" + HueBridgeParameterType;
        URL url1 = new URL(finalURL);
        connection = (HttpURLConnection) url1.openConnection();
        connection.connect();
        InputStream stream1 = connection.getInputStream();
        BufferedReader reader1 = new BufferedReader(new InputStreamReader(stream1));
        StringBuffer br1 = new StringBuffer();
        String line1 = " ";
        while ((line1 = reader1.readLine()) != null) {
            br1.append(line1);
        }

        String output2 = br1.toString();
        JSONObject jsonObject2 = new JSONObject(output2);

        Object ob2 = jsonObject2.get("action");
        String newString2 = ob2.toString();
        JSONObject jsonObject3 = new JSONObject(newString2);
        Object ob3 = jsonObject3.get("on");
        System.out.println(ob3);


        if (ob3.toString().equals("true"))

        {
            Status = "1";
            ActualResult = "Group " + RoomName + " is turned on";
            Comments = "NA";
            ExpectedResult= "Group " + RoomName + " should be turned on";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        } else {
            Status = "0";
            ActualResult = "Group " + RoomName + " is not turned on";
            Comments = "FAIL: Group Status of " + RoomName + " is : " + ob3.toString();
            ExpectedResult= "Group " + RoomName + " should turned on";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }
        //CALLING THE FUNCTION FOR WRITING THE CODE IN EXCEL FILE
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);
    }
    //WRITING THE RESULT IN EXCEL FILE
    public String CurrentdateTime;
    public int nextRowNumber;
    public void storeResultsExcel (String excelStatus, String excelActualResult, String excelComments, String resultFileName, String ExcelExpectedResult, String resultAPIVersion, String resultSWVersion) throws IOException
    {
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
        r2c2.setCellValue("15");

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
