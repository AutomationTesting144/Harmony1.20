package com.example.a310287808.harmony120;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.File;
import java.io.FileOutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Example local unit test, which will execute on the development machine (host).
 *
 * @see <a href="http://d.android.com/tools/testing">Testing documentation</a>
 */
public class ControllingScript {
    public String filename;
    public int Month;
    public int Date;
    AndroidDriver driver;
    Dimension size;

    @Before
    public void setUp() throws MalformedURLException {
        DesiredCapabilities capabilities = new DesiredCapabilities();
        capabilities.setCapability("deviceName", "Nexus 6");
        capabilities.setCapability(CapabilityType.BROWSER_NAME, "Android");
        capabilities.setCapability(CapabilityType.VERSION, "7.0");
        capabilities.setCapability("platformName", "Android");
        capabilities.setCapability("appPackage", "com.google.android.calculator");
        capabilities.setCapability("appActivity", "com.android.calculator2.Calculator");
        capabilities.setCapability("newCommandTimeout", "45000");


        driver = new AndroidDriver(new URL("http://localhost:4723/wd/hub"), capabilities);
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

    }

    public String APIVersion;
    public String SWVersion;
    public String Currentdate;
    public int fileCounter = 1;
    public HSSFWorkbook workbook;
    public HSSFSheet sheet1;
    public String resultFileName;

    @Test
    public void testRuns() throws Exception {
        File ArchiveFolder = new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\Archive");
        File OldFile = new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\");
        File[] oldFiles = OldFile.listFiles();

        Calendar cal = Calendar.getInstance();
        SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
        Currentdate = sdf.format(cal.getTime());

        for (int i = 0; i < oldFiles.length; i++) {
            //System.out.println(oldFiles[i]);
            Boolean status = oldFiles[i].getName().contains("HarmonyAutomationResults");
            //System.out.println(status);
            if (status == true) {
                fileCounter = 0;
                String filename1 = oldFiles[i].getName();
                if (filename1.contains(Currentdate)) {

                    break;
                } else {
                    oldFiles[i].renameTo(new File(ArchiveFolder + "\\" + oldFiles[i].getName()));
                    filename = "C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\HarmonyAutomationResults-" + Currentdate + ".xls";
                    //System.out.println(filename);
                    workbook = new HSSFWorkbook();
                    sheet1 = workbook.createSheet(sdf.format(cal.getTime()));
                    HSSFRow rowhead = sheet1.createRow((short) 0);
                    rowhead.createCell((short) 2).setCellValue(new HSSFRichTextString("isPassed"));
                    rowhead.createCell((short) 1).setCellValue(new HSSFRichTextString("Test Case Id"));
                    rowhead.createCell((short) 0).setCellValue(new HSSFRichTextString("RunDateTime"));
                    rowhead.createCell((short) 3).setCellValue(new HSSFRichTextString("Actual Result"));
                    rowhead.createCell((short) 4).setCellValue(new HSSFRichTextString("Failure Reason"));
                    rowhead.createCell((short) 5).setCellValue(new HSSFRichTextString("API Version"));
                    rowhead.createCell((short) 6).setCellValue(new HSSFRichTextString("SW Version"));
                    FileOutputStream fileOut = new FileOutputStream(filename);
                    workbook.write(fileOut);
                    fileOut.close();

                }

            }

        }

        if (fileCounter != 0) {
            //FileInputStream fsIP= new FileInputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\"+resultFileName));
            filename = "C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\HarmonyAutomationResults-" + Currentdate + ".xls";
            //System.out.println(filename);
            workbook = new HSSFWorkbook();
            sheet1 = workbook.createSheet(sdf.format(cal.getTime()));
            HSSFRow rowhead = sheet1.createRow((short) 0);
            rowhead.createCell((short) 2).setCellValue(new HSSFRichTextString("isPassed"));
            rowhead.createCell((short) 1).setCellValue(new HSSFRichTextString("Test Case Id"));
            rowhead.createCell((short) 0).setCellValue(new HSSFRichTextString("RunDateTime"));
            rowhead.createCell((short) 3).setCellValue(new HSSFRichTextString("Actual Result"));
            rowhead.createCell((short) 4).setCellValue(new HSSFRichTextString("Failure Reason"));
            rowhead.createCell((short) 5).setCellValue(new HSSFRichTextString("API Version"));
            rowhead.createCell((short) 6).setCellValue(new HSSFRichTextString("SW Version"));
            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
        }


        File newFile = new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\");
        File[] newFiles = newFile.listFiles();

        for (int i = 0; i < newFiles.length; i++) {
            Boolean fileYesNo = newFiles[i].getName().contains("HarmonyAutomationResults");

            if (fileYesNo == true) {

                resultFileName = newFiles[i].getName();
            }
        }


        APIVersion apiv = new APIVersion();
        APIVersion= apiv.getAPIVersion();


        SWVersion swv = new SWVersion();
        SWVersion = swv.getSWVersion();

//        BridgeAuthenticationFirstTime Bridge=new BridgeAuthenticationFirstTime();
//        Bridge.BridgeAuthenticationFirstTime(driver, resultFileName,APIVersion,SWVersion);

//        NewWhiteList  WhiteList=new NewWhiteList();
//        WhiteList.NewWhiteList(driver,resultFileName,APIVersion,SWVersion);
//
//        WhiteListFriendlyName  FriendlyName=new WhiteListFriendlyName();
//        FriendlyName.WhiteListFriendlyName(driver,resultFileName,APIVersion,SWVersion);

//        BridgeConnectionAfterWhiteListDeletion AfterWhiteListDeletion=new BridgeConnectionAfterWhiteListDeletion();
//        AfterWhiteListDeletion.BridgeConnectionAfterWhiteListDeletion(driver, resultFileName,APIVersion,SWVersion);

        TurnOnLight OnLight=new TurnOnLight();
        OnLight.TurnOnLight(driver, resultFileName,APIVersion,SWVersion);

        TurnOffLight OffLight=new TurnOffLight();
        OffLight.TurnOffLight(driver, resultFileName,APIVersion,SWVersion);

        SingleLightBrightness LightBrightness=new SingleLightBrightness();
        LightBrightness.SingleLightBrightness(driver, resultFileName,APIVersion,SWVersion);

        ColorChangeSingle colorSingle=new ColorChangeSingle();
        colorSingle.ColorChangeSingle(driver, resultFileName,APIVersion,SWVersion);

        LightAddition addition=new LightAddition();
        addition.LightAddition(driver, resultFileName,APIVersion,SWVersion);

        LightRename rename=new LightRename();
        rename.LightRename(driver, resultFileName,APIVersion,SWVersion);

        LightDelete delete=new LightDelete();
        delete.LightDelete(driver, resultFileName,APIVersion,SWVersion);
//
        GroupCreation creation=new GroupCreation();
        creation.GroupCreation(driver, resultFileName,APIVersion,SWVersion);

        GroupOn On=new GroupOn();
        On.GroupOn(driver, resultFileName,APIVersion,SWVersion);

        GroupOff Off=new GroupOff();
        Off.GroupOff(driver, resultFileName,APIVersion,SWVersion);
//
        GroupRename rename1=new GroupRename();
        rename1.GroupRename(driver, resultFileName,APIVersion,SWVersion);

        GroupDeletion delete1=new GroupDeletion();
        delete1.GroupDeletion(driver, resultFileName,APIVersion,SWVersion);

        ActivityCreation creation1=new ActivityCreation();
        creation1.ActivityCreation(driver, resultFileName,APIVersion,SWVersion);
//
        ActivityON ON=new ActivityON();
        ON.ActivityON(driver, resultFileName,APIVersion,SWVersion);

        ActivityOFF OFF=new ActivityOFF();
        OFF.ActivityOFF(driver, resultFileName,APIVersion,SWVersion);

        ActivityChange change=new ActivityChange();
        change.ActivityChange(driver, resultFileName,APIVersion,SWVersion);

        ActivityDelete delete2=new ActivityDelete();
        delete2.ActivityDelete(driver, resultFileName,APIVersion,SWVersion);

//        WhiteListDeletion  Deletion=new WhiteListDeletion();
//        Deletion.WhiteListDeletion(driver);

    }
}