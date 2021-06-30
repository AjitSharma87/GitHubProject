package main;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class ReadTestData {

    // Static Data Class
    static class TestDataDefined{
        String ResultToALM;
        String Property;
        String Action;
        String Type;
        String Value;
        String Screenshot;
        String Description;
        String ExpectedResult;
        String PassResult;
        String FailResult;
        String Status;

        @Override
        public String toString() {
            return "TestDataDefined{" +
                    "ResultToALM='" + ResultToALM + '\'' +
                    ", Property='" + Property + '\'' +
                    ", Type='" + Action + '\'' +
                    ", Function='" + Type + '\'' +
                    ", Value='" + Value + '\'' +
                    ", Screenshot='" + Screenshot + '\'' +
                    ", Description='" + Description + '\'' +
                    ", ExpectedResult='" + ExpectedResult + '\'' +
                    ", PassResult='" + PassResult + '\'' +
                    ", FailResult='" + FailResult + '\'' +
                    ", Status='" + Status + '\'' +
                    '}';
        }
    }
    static FnLib functions = new FnLib();
    static QcBridge qcBridge = new QcBridge();

    public static void openExcel(){
        try {
            // Initialized TestData File
            File file = new File("D:\\pocExcelFramework\\TestData\\TestData.xlsx");

            // Throws IOException, InvalidFormatException
            Workbook workbook = new XSSFWorkbook(file);

            // Iterate through all sheets
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);

                ArrayList<TestDataDefined> TestData = getTestData(sheet);
                functions.addScenario(sheet.getSheetName());

                // Iterate through all steps
                for (TestDataDefined t :TestData) {
                    // Perform step
                    t.Status = doStep(t);
                    if (t.ResultToALM.equalsIgnoreCase("Y") && !t.Screenshot.equalsIgnoreCase("Y")){
                        if (t.Status.equalsIgnoreCase("Passed")){
                            qcBridge.updateSteps(t.Status,t.Description,t.ExpectedResult,t.PassResult);
                        } else{
                            qcBridge.updateSteps(t.Status,t.Description,t.ExpectedResult,t.FailResult);
                        }
                    }

                    if (t.ResultToALM.equalsIgnoreCase("Y") && t.Screenshot.equalsIgnoreCase("Y")){
                        if (t.Status.equalsIgnoreCase("Passed")){
                            qcBridge.updateStepsWithScreenshot(t.Status,t.Description,t.ExpectedResult,t.PassResult);
                        } else{
                            qcBridge.updateStepsWithScreenshot(t.Status,t.Description,t.ExpectedResult,t.FailResult);
                        }
                    }
//                    System.out.println(t.toString());

                }

            }

        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static String doStep(TestDataDefined t) {
        String stepResult = "FAILED";
        switch (t.Action){
            case "TypeToElement":{
                try{
                    functions.sendKeysToElem(t.Type,t.Property,t.Value);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "ClickElement":{
                try{
                    functions.clickElem(t.Type,t.Property);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "HoverElement":{
                try{
                    functions.hoverOverElem(t.Type,t.Property);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "VerifyElementIsVisible":{
                try{
                    functions.isExistsElem(t.Type,t.Property);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "CloseExe":{
                try{
                    Runtime rt = Runtime.getRuntime();
                    functions.closeExe(t.Value);
                    return "Passed";
                }catch (Exception e){
                    e.printStackTrace();
                    return "Failed";
                }
            }
            case "ClearResultsFolder":{
                try {
                    functions.clearResultsFolder(t.Value);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "CreateExcelResult":{
                try {
                    functions.CreateExcelReport(t.Value);
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }
            case "Other":{
                switch (t.Type){
                    case "SendKeys":{
                        try {
                            functions.send_keys(t.Value,'N');
                            return "Passed";
                        }catch (Exception e){
                            return "Failed";
                        }
                    }
                    case "OpenUrl": {
                        try{
                            functions.openUrl(t.Value);
                            return "Passed";
                        }catch (Exception e){
//                            e.printStackTrace();
                            return "Failed";
                        }
                    }

                }
            }
            case "Wait":{
                try{
                    functions.addWait(Integer.parseInt(t.Value));
                    return "Passed";
                }catch (Exception e){
                    return "Failed";
                }
            }

        }
        return stepResult;
    }

    public static ArrayList<TestDataDefined> getTestData(Sheet sheet){
        ArrayList<TestDataDefined> mappedTestData = new ArrayList<>();
        // Skip if test data definition else execute test case
        if (!sheet.getSheetName().equals("TestDataDefined_DO_NOT_CHANGE") && !sheet.getSheetName().equals("ALM_DETAILS")){
            // Iterate through all rows
            for (Row r: sheet){
                TestDataDefined newData = new TestDataDefined();
                if (r.getRowNum() != 0){
                    // Iterate through all cells
                    for (Cell cell:r)
                    {
                        int colIndex = cell.getColumnIndex();
                        String colValue =  String.valueOf(cell);
                        switch (colIndex){
                            case 0: { newData.ResultToALM = colValue; break; }
                            case 1: { newData.Property = colValue; break; }
                            case 2: { newData.Action = colValue; break; }
                            case 3: { newData.Type = colValue; break;}
                            case 4: { newData.Value = colValue; break; }
                            case 5: { newData.Screenshot = colValue; break; }
                            case 6: { newData.Description = colValue; break; }
                            case 7: { newData.ExpectedResult = colValue; break; }
                            case 8: { newData.PassResult = colValue; break; }
                            case 9: { newData.FailResult = colValue; break; }
                        }
                    }
                    mappedTestData.add(newData);
                }
            }
        }
        return mappedTestData;
    }

    public static void main(String[] args) {
        qcBridge.openConnection();
        qcBridge.createRun();
        openExcel();

    }




}