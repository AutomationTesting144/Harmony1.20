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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 8/2/2017.
 */

public class ColorChangeSingle {
    public String IPAddress = "192.168.86.23/api";
    public String HueUserName = "SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2";
    public String HueBridgeParameterType = "lights/6";
    public String finalURL;
    public String lightStatusReturned;
    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;

    AndroidDriver driver;

    public  void ColorChangeSingle(AndroidDriver driver,String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException, ParseException {
        driver.navigate().back();
        driver.navigate().back();
        HttpURLConnection connection;

        //Checking whether the group light is ON/OFF
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

        //If the lights in the group are already ON then turn them off
        if (ob1.toString()==("true"))
        {
            URL url1 = new URL("http://192.168.86.23/api/SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2/lights/6/state");
            String content = "{"+"\"on\""+":"+"false"+"}";
            HttpURLConnection httpCon = (HttpURLConnection) url1.openConnection();
            httpCon.setDoOutput(true);
            httpCon.setRequestMethod("PUT");
            OutputStreamWriter out = new OutputStreamWriter(httpCon.getOutputStream());
            out.write(content);
            out.close();
            httpCon.getInputStream();
            System.out.println(httpCon.getResponseCode());
            TimeUnit.SECONDS.sleep(5);
            System.out.println("Lights are switched off");
            TimeUnit.SECONDS.sleep(5);
        }

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Harmony']")).click();
        TimeUnit.SECONDS.sleep(10);

        //clicking on device
        driver.findElement(By.id("com.logitech.harmonyhub:id/home_tab_devices")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on downward arrow
        driver.findElement(By.id("com.logitech.harmonyhub:id/device_right_icon")).click();
        TimeUnit.SECONDS.sleep(2);

        //clicking on bulb icon
        driver.findElement(By.xpath("//android.widget.ImageView[@bounds='[78,1082][222,1226]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //choosing a light to decrease brightness
        driver.findElement(By.xpath("//android.widget.ImageView[@bounds='[1265,1101][1370,1206]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //change the color of light
        URL url1 = new URL("http://192.168.86.23/api/SP8cBJNKikQDk9e2AvtZEoL55N63FhVVFSwceVM2/lights/6/state");
        String content = "{"+"\"xy\""+":"+"[0.4090,0.5180]"+"}";
        HttpURLConnection httpCon = (HttpURLConnection) url1.openConnection();
        httpCon.setDoOutput(true);
        httpCon.setRequestMethod("PUT");
        OutputStreamWriter out = new OutputStreamWriter(httpCon.getOutputStream());
        out.write(content);
        out.close();
        httpCon.getInputStream();
        System.out.println(httpCon.getResponseCode());
        TimeUnit.SECONDS.sleep(5);
        System.out.println("color changed for the light");
        TimeUnit.SECONDS.sleep(5);

        //choosing color
        driver.findElement(By.xpath("//android.widget.RelativeLayout[@bounds='[53,442][1440,2030]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Fetching the result from API
        finalURL = "http://" + IPAddress + "/" + HueUserName + "/" + HueBridgeParameterType;
        URL url4 = new URL(finalURL);
        connection = (HttpURLConnection) url4.openConnection();
        connection.connect();
        InputStream stream4 = connection.getInputStream();
        BufferedReader reader4 = new BufferedReader(new InputStreamReader(stream4));
        br = new StringBuffer();
        String line4 = " ";
        while ((line4 = reader4.readLine()) != null) {
            br.append(line4);
        }
        String output = br.toString();
        ColorChangeSingleStatus SingleStatus = new ColorChangeSingleStatus();
        lightStatusReturned = SingleStatus.ColorChangeSingleStatus(output);

        String Xval=lightStatusReturned.substring(1,5);
        String Yval=lightStatusReturned.substring(8,12);
        String Xgreen="0.47";
        String Ygreen="0.42";

        String output2 = br.toString();
        JSONObject jsonObject2 = new JSONObject(output2);

        Object ob2 = jsonObject2.get("state");
//        String newString2 = ob2.toString();
        Object lightNameObject = jsonObject2.get("name");
        String lightName = lightNameObject.toString();

        br.append(lightName);
        br.append("\n");

        if ((Xval.equals(Xgreen)) && (Yval.equals(Ygreen))) {
            Status = "1";
            ActualResult ="Color changed for the selected light"+lightName;
            Comments = "NA";
            ExpectedResult= " Light: "+lightName+" should change the color.";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        } else {
            Status = "0";
            ActualResult ="Color does not changed for the selected light"+"\n"+lightName;
            Comments = "FAIL:color is not changed for the LIGHT";
            ExpectedResult= " Light: "+lightName+" should change the color.";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        driver.navigate().back();
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
        r2c2.setCellValue("8");

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
