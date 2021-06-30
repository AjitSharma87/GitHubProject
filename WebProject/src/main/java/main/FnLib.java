package main;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;


public class FnLib {

    private static WebDriver driver;
    private static int stepCount = 0;
    private static int scrShot = 0;
    private static String scenarioName = "";
    private static int failureStatus = 0;
//    static {System.setProperty("webdriver.chrome.driver","C:\\drivers\\chromedriver.exe");	}
    //static {System.setProperty("webdriver.ie.driver","C:\\drivers\\IEDriverServer.exe");	}


    FnLib(){
        File chromeDriver = new File("C:\\drivers\\chromedriver.exe");
        System.setProperty("webdriver.chrome.driver", chromeDriver.getAbsolutePath());
        driver = new ChromeDriver();
//        System.setProperty("webdriver.chrome.driver","C:\\drivers\\chromedriver.exe");
//        driver = new ChromeDriver();
    }

    static class reporter{
        String StepName,Status,ExpectedResult,ActualResult;
        reporter(String StepName,String Status,String ExpectedResult,String ActualResult){
            this.StepName = StepName;
            this.Status = Status;
            this.ExpectedResult = ExpectedResult;
            this.ActualResult = ActualResult;
        }
    }

    public static void addScenario(String sc) {
        scenarioName = sc;
        stepCount = 0;
        scrShot = 0;
        failureStatus = 0;
        reportList.clear();
//        clearResultsFolder("C:/demoWebTest/Results/"+scenarioName);
//        System.out.println(sc);
    }

    public static void clearResultsFolder(String path) {

        File index = new File(path);
        if(index.exists()) {
            String[]entries = index.list();
            for(String s: entries){
                File currentFile = new File(index.getPath(),s);
                currentFile.delete();
            }
        }
    }

    public static  ArrayList<reporter> reportList=new ArrayList<reporter>();

    public static void report( String status, String expectedResult, String actualResult) {
        stepCount++;
        reporter step = new reporter("Step"+stepCount,status,expectedResult,actualResult+" \n Refer to Step"+stepCount+".png");
        reportList.add(step);
        takeScreenshot(stepCount);
    }

    public static void CreateExcelReport(String path) throws IOException {
        String[] columns = {"Step Name", "Status", "Expected Result", "Actual Result"};

        Workbook workbook = new XSSFWorkbook();
        // Create a Sheet
        Sheet sheet = workbook.createSheet(scenarioName);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        CellStyle failCellStyle = workbook.createCellStyle();
        Font failFont = workbook.createFont();
        failFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        failCellStyle.setFont(failFont);

        CellStyle passCellStyle = workbook.createCellStyle();
        Font passFont = workbook.createFont();
        passFont.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        passCellStyle.setFont(passFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        // Create Other rows and cells with employees data
        int rowNum = 1;
        Iterator<reporter> itr= reportList.iterator();
        //traversing elements of ArrayList object
        while(itr.hasNext()){
            Row row = sheet.createRow(rowNum++);
            reporter st=(reporter)itr.next();
            //System.out.println(st.StepName+" -- "+st.Status+" -- "+st.ExpectedResult+" -- "+st.ActualResult);

            if (st.Status == "Passed") {
                row.createCell(0).setCellValue(st.StepName);
                row.getCell(0).setCellStyle(passCellStyle);
                row.createCell(1).setCellValue(st.Status);
                row.getCell(1).setCellStyle(passCellStyle);
                row.createCell(2).setCellValue(st.ExpectedResult);
                row.getCell(2).setCellStyle(passCellStyle);
                row.createCell(3).setCellValue(st.ActualResult);
                row.getCell(3).setCellStyle(passCellStyle);
            }
            if (st.Status == "Failed") {
                row.createCell(0).setCellValue(st.StepName);
                row.getCell(0).setCellStyle(failCellStyle);
                row.createCell(1).setCellValue(st.Status);
                row.getCell(1).setCellStyle(failCellStyle);
                row.createCell(2).setCellValue(st.ExpectedResult);
                row.getCell(2).setCellStyle(failCellStyle);
                row.createCell(3).setCellValue(st.ActualResult);
                row.getCell(3).setCellStyle(failCellStyle);
            }

        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream(path+scenarioName+"/Result.xlsx");

        // Write the output to a file

        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    public static void printReportList() {
        Iterator<reporter> itr= reportList.iterator();
        //traversing elements of ArrayList object
        while(itr.hasNext()){
            reporter st=(reporter)itr.next();
            System.out.println(st.StepName+" -- "+st.Status+" -- "+st.ExpectedResult+" -- "+st.ActualResult);
        }
    }

    public static void highLightElem(WebElement element){
        JavascriptExecutor js = (JavascriptExecutor)driver;
        js.executeScript("arguments[0].setAttribute('style','background: yellow; border: 2px solid red;');", element);
        try {
            Thread.sleep(400);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        js.executeScript("arguments[0].setAttribute('style','background: none; border: none;');", element);
        try {
            Thread.sleep(200);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        js.executeScript("arguments[0].setAttribute('style','background: yellow; border: 2px solid red;');", element);
        try {
            Thread.sleep(200);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        js.executeScript("arguments[0].setAttribute('style','background: none; border: none;');", element);
        try {
            Thread.sleep(100);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }


    public static void launchBrowser() {
        ChromeOptions options = new ChromeOptions();
        options.addArguments("disable-infobars");
        options.addArguments("start-maximized");
        driver = new ChromeDriver(options);
        //driver = new InternetExplorerDriver();
    }

    public static void closeBrowser() {
        driver.close();
    }

    public static void closeExe(String exe) throws IOException {
        Runtime.getRuntime().exec("taskkill /F /IM "+exe+".exe");

    }

    public static void openUrl(String url) throws InterruptedException {
        driver.get("https://"+url);
        Thread.sleep(4000);
    }

    public static File takeScreenshot(int stepCount) {
        File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
        try {
            // now copy the  screenshot to desired location using copyFile //method
            FileUtils.copyFile(src, new File("D:/pocExcelFramework/Results/"+scenarioName+"/Step"+stepCount+".png"));

        }

        catch (IOException e)
        {
            System.out.println(e.getMessage());
        }
        File f = new File("D:/pocExcelFramework/Results/"+scenarioName+"/Step"+stepCount+".png");
        return f;
    }

    public static void onlyScreenshot() {
        File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
        try {
            // now copy the  screenshot to desired location using copyFile //method
            FileUtils.copyFile(src, new File("C:/demoWebTest/Results/"+scenarioName+"/ScreenShot"+scrShot+".png"));
            scrShot++;

        }

        catch (IOException e)
        {
            System.out.println(e.getMessage());

        }
    }

    public static void clickElem(String objProp, String obj) {
        switch(objProp.toLowerCase()) {
            case "xpath":
                if(driver.findElement(By.xpath(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.xpath(obj)));
                    driver.findElement(By.xpath(obj)).click();
                    //report("Passed",obj+" Button should be clicked",obj+" Button is clicked successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    //report("Failed",obj+" Button should be clicked",obj+" Button is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "name":
                if(driver.findElement(By.name(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.name(obj)));
                    driver.findElement(By.name(obj)).click();
                    //report("Passed",obj+" Button should be clicked",obj+" Button is clicked successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    //report("Failed",obj+" Button should be clicked",obj+" Button is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "id":
                if(driver.findElement(By.id(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.id(obj)));
                    driver.findElement(By.id(obj)).click();
                    //report("Passed",obj+" Button should be clicked",obj+" Button is clicked successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    //report("Failed",obj+" Button should be clicked",obj+" Button is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            default:
                if(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).isDisplayed()) {
                    highLightElem(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")));
                    driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).click();
                    //report("Passed",obj+" Button should be clicked",obj+" Button is clicked successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    //report("Failed",obj+" Button should be clicked",obj+" Button is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
        }
    }

    public static void sendKeysToElem(String objProp, String obj, String val) {
        switch(objProp.toLowerCase()) {
            case "xpath":
                if(driver.findElement(By.xpath(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.xpath(obj)));
                    WebElement elem = driver.findElement(By.xpath(obj));
                    elem.sendKeys(val);
                    //report("Passed",obj+" Element should set to value" + val,val+" value is inserted successfully");
                    //System.out.println(obj+" value is inserted");
                } else {
                    //report("Failed",obj+" Element should set to value" + val,obj+" element is NOT displayed");
                    //System.out.println(obj+" value is NOT displayed");
                }
                break;
            case "name":
                if(driver.findElement(By.name(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.name(obj)));
                    WebElement elem = driver.findElement(By.name(obj));
                    elem.sendKeys(val);
                    //report("Passed",obj+" Element should set to value","value is inserted successfully");
                    //System.out.println(obj+" value is inserted");
                } else {
                    //report("Failed",obj+" Element should set to value",obj+" element is NOT displayed");
                    //System.out.println(obj+" value is NOT displayed");
                }
                break;
            case "id":
                if(driver.findElement(By.id(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.id(obj)));
                    WebElement elem = driver.findElement(By.id(obj));
                    elem.sendKeys(val);
                    //report("Passed",obj+" Element should set to value" ,"value is inserted successfully");
                    //System.out.println(obj+" value is inserted");
                } else {
                    //report("Failed",obj+" Element should set to value" ,obj+" element is NOT displayed");
                    //System.out.println(obj+" is NOT displayed");
                }
                break;
            default:
                if(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).isDisplayed()) {
                    highLightElem(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")));
                    WebElement elem = driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']"));
                    elem.sendKeys(val);
                    //report("Passed",obj+" Element should set to value","value is inserted successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    //report("Failed",obj+" Element should set to value",obj+" element is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
        }
    }

    public static void isExistsElem(String objProp, String obj) {
        switch(objProp.toLowerCase()) {
            case "xpath":
                if(driver.findElement(By.xpath(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.xpath(obj)));
//                    report("Passed",name+" should exist",name+" is displayed successfully");
                    //System.out.println(obj+" exists");
                } else {
//                    report("Failed",name+" should exist",name+" is NOT displayed");
                    //System.out.println(obj+" is NOT displayed");
                }
                break;
            case "name":
                if(driver.findElement(By.name(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.name(obj)));
//                    report("Passed",name+" Element should exist",name+" is displayed successfully");
                    //System.out.println(obj+" exists");
                } else {
//                    report("Failed",name+" Element should exist",name+" is NOT displayed");
                    //System.out.println(obj+" is NOT displayed");
                }
                break;
            case "id":
                if(driver.findElement(By.id(obj)).isDisplayed()) {
                    highLightElem(driver.findElement(By.id(obj)));
//                    report("Passed",name+" Element should exist",name+" is displayed successfully");
                    //System.out.println(obj+" exists");
                } else {
//                    report("Failed",name+" Element should exist",name+" is NOT displayed");
                    //System.out.println(obj+" is NOT displayed");
                }
                break;
            default:
                if(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).isDisplayed()) {
                    highLightElem(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")));
//                    report("Passed",name+" Element should exist",name+" is displayed successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
//                    report("Failed",name+" Element should exist",name+" is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
        }
    }

    public static void addWait(int sec) throws InterruptedException {
        sec = (sec*1000);
        Thread.sleep(sec);
    }

    public static void hoverOverElem(String objProp, String obj) {
        switch(objProp.toLowerCase()) {
            case "xpath":
                if(driver.findElement(By.xpath(obj)).isDisplayed()) {
                    Actions actions = new Actions(driver);
                    highLightElem(driver.findElement(By.xpath(obj)));
                    WebElement elem = driver.findElement(By.xpath(obj));
                    actions.moveToElement(elem).perform();
                    report("Passed",obj+" Element should Hover over",obj+" Element is hovered successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed",obj+" Element should Hover over",obj+" Element is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "name":
                if(driver.findElement(By.name(obj)).isDisplayed()) {
                    Actions actions = new Actions(driver);
                    highLightElem(driver.findElement(By.name(obj)));
                    WebElement elem = driver.findElement(By.name(obj));
                    actions.moveToElement(elem).perform();
                    report("Passed",obj+" Element should Hover over",obj+" Element is hovered successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed",obj+" Element should Hover over",obj+" Element is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "id":
                if(driver.findElement(By.id(obj)).isDisplayed()) {
                    Actions actions = new Actions(driver);
                    highLightElem(driver.findElement(By.id(obj)));
                    WebElement elem = driver.findElement(By.id(obj));
                    actions.moveToElement(elem).perform();
                    report("Passed",obj+" Element should Hover over",obj+" Element is hovered successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed",obj+" Element should Hover over",obj+" Element is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            default:
                if(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).isDisplayed()) {
                    Actions actions = new Actions(driver);
                    highLightElem(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")));
                    WebElement elem =  driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']"));
                    actions.moveToElement(elem).perform();
                    report("Passed",obj+" Element should Hover over",obj+" Element is hovered successfully");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed",obj+" Element should Hover over",obj+" Element is NOT displayed");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
        }
    }

    public static String PasswordDecryption(String EncryptedString) {
        char intPosition;
        int intCounter;
        String strTemp = "";
        StringBuilder sb = new StringBuilder(EncryptedString);
        EncryptedString = (sb).reverse().toString();
        for (intCounter = 1; (intCounter <= EncryptedString.length()); intCounter++) {
            intPosition = EncryptedString.charAt(intCounter-1);
            int ascii = (int)intPosition;
            strTemp = (strTemp + ((char)((ascii - 1))));
        }

        return strTemp;
    }

    public static void CountLinks(String objProp, String obj) {
        switch(objProp.toLowerCase()) {
            case "tagname":
                if(driver.findElement(By.tagName(obj)).isDisplayed()) {
                    java.util.List<WebElement> links = driver.findElements(By.tagName(obj));
                    for (int i = 1; i<=links.size(); i=i+1) {
                        highLightElem(links.get(i));
                    }
                    report("Passed","Count Number of links",links.size()+" Links present");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed","Count Number of links","NO links found");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "xpath":
                if(driver.findElement(By.xpath(obj)).isDisplayed()) {
                    java.util.List<WebElement> links = driver.findElements(By.xpath(obj));
                    for (int i = 1; i<=links.size(); i=i+1) {
                        highLightElem(links.get(i));
                    }
                    report("Passed","Count Number of links",links.size()+" Links present");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed","Count Number of links","NO links found");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "name":
                if(driver.findElement(By.name(obj)).isDisplayed()) {
                    java.util.List<WebElement> links = driver.findElements(By.name(obj));
                    for (int i = 1; i<=links.size(); i=i+1) {
                        highLightElem(links.get(i));
                    }
                    report("Passed","Count Number of links",links.size()+" Links present");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed","Count Number of links","NO links found");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            case "id":
                if(driver.findElement(By.id(obj)).isDisplayed()) {
                    java.util.List<WebElement> links = driver.findElements(By.id(obj));
                    for (int i = 1; i<=links.size(); i=i+1) {
                        highLightElem(links.get(i));
                    }
                    report("Passed","Count Number of links",links.size()+" Links present");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed","Count Number of links","NO links found");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
            default:
                if(driver.findElement(By.cssSelector("a["+objProp+"='"+obj+"']")).isDisplayed()) {
                    java.util.List<WebElement> links = driver.findElements(By.cssSelector("a["+objProp+"='"+obj+"']"));
                    for (int i = 1; i<=links.size(); i=i+1) {
                        highLightElem(links.get(i));
                    }
                    report("Passed","Count Number of links",links.size()+" Links present");
                    //System.out.println(obj+" Button is clicked");
                } else {
                    report("Failed","Count Number of links","NO links found");
                    //System.out.println(obj+" Button is NOT displayed");
                }
                break;
        }
    }

    public static void send_keys(String str, char flag) throws AWTException {
        String text = str;
        if(flag == 'Y') {
            text = PasswordDecryption(str);
        }
        StringSelection stringSelection = new StringSelection(text);
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        clipboard.setContents(stringSelection, stringSelection);

        Robot robot = new Robot();
        robot.keyPress(KeyEvent.VK_CONTROL);
        robot.keyPress(KeyEvent.VK_V);
        robot.keyRelease(KeyEvent.VK_V);
        robot.keyRelease(KeyEvent.VK_CONTROL);
    }


    public static void send_tab() throws AWTException {
        Robot robot = new Robot();
        // Simulate key Events
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_TAB);
    }
    public static void send_enter() throws AWTException {
        Robot robot = new Robot();
        // Simulate key Events
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
    }



}
