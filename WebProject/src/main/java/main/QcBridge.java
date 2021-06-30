package main;
import com4j.Com4jObject;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ota.ClassFactory;
import ota.IAttachment;
import ota.IAttachmentFactory;
import ota.IBaseFactory;
import ota.IExtendedStorage;
import ota.IList;
import ota.IRun;
import ota.IRunFactory;
import ota.IStep;
import ota.ITDConnection;
import ota.ITSTest;
import ota.ITestSet;
import ota.ITestSetFolder;
import ota.ITestSetTreeManager;

public class QcBridge {
    private static final String url = "http://alm1152.pfizer.com:8080/qcbin";
    private static String username;
    private static String password;
    private static String domain;
    private static String project;
    private static String testName;
    private static String testSetName;
    private static String testSetFolder;
    private static ITDConnection tdConnection;
    private static IRun currentRun = null;
    private static int stepCount = 0;

    QcBridge() {
        try {
            File file = new File("D:\\pocExcelFramework\\TestData\\TestData.xlsx");
            Workbook workbook = new XSSFWorkbook(file);

            label52:
            for(int i = 0; i < workbook.getNumberOfSheets(); ++i) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet.getSheetName().equals("ALM_DETAILS")) {
                    Iterator var5 = sheet.iterator();

                    while(true) {
                        Row r;
                        do {
                            if (!var5.hasNext()) {
                                continue label52;
                            }

                            r = (Row)var5.next();
                        } while(r.getRowNum() != 1);

                        Iterator var7 = r.iterator();

                        while(var7.hasNext()) {
                            Cell cell = (Cell)var7.next();
                            int colIndex = cell.getColumnIndex();
                            String colValue = String.valueOf(cell);
                            switch(colIndex) {
                                case 0:
                                    username = colValue;
                                    break;
                                case 1:
                                    password = FnLib.PasswordDecryption(colValue);
                                    break;
                                case 2:
                                    domain = colValue;
                                    break;
                                case 3:
                                    project = colValue;
                                    break;
                                case 4:
                                    testName = colValue;
                                    break;
                                case 5:
                                    testSetName = colValue;
                                    break;
                                case 6:
                                    testSetFolder = colValue;
                            }
                        }
                    }
                }
            }
        } catch (InvalidFormatException | IOException var11) {
            var11.printStackTrace();
        }

        tdConnection = ClassFactory.createTDConnection();
    }

    public static void openConnection() {
        tdConnection.initConnectionEx("http://alm1152.pfizer.com:8080/qcbin");
        tdConnection.connectProjectEx(domain, project, username, password);
    }

    public static void createRun() {
        if (tdConnection.connected()) {
            ITestSetTreeManager testSetTreeManager = (ITestSetTreeManager)tdConnection.testSetTreeManager().queryInterface(ITestSetTreeManager.class);
            ITestSetFolder testSetFolder = (ITestSetFolder)testSetTreeManager.nodeByPath(QcBridge.testSetFolder).queryInterface(ITestSetFolder.class);
            IList testSetList = (IList)testSetFolder.findTestSets(testSetName, false, (String)null).queryInterface(IList.class);
            Iterator var3 = testSetList.iterator();

            while(var3.hasNext()) {
                Com4jObject tSetIterator = (Com4jObject)var3.next();
                ITestSet testSet = (ITestSet)tSetIterator.queryInterface(ITestSet.class);
                IBaseFactory testFactory = (IBaseFactory)testSet.tsTestFactory().queryInterface(IBaseFactory.class);
                IList testInstances = testFactory.newList("");
                Iterator var8 = testInstances.iterator();

                while(var8.hasNext()) {
                    Com4jObject testInstanceObj = (Com4jObject)var8.next();
                    ITSTest testInstance = (ITSTest)testInstanceObj.queryInterface(ITSTest.class);
                    if (testInstance.testName().equals(testName)) {
                        IRunFactory runFactory = (IRunFactory)testInstance.runFactory().queryInterface(IRunFactory.class);
                        IRun newRun = (IRun)runFactory.addItem("Java Run at " + String.format("%1$tm/%1$td/%1$tY %1$tH:%1$tM:%1$tS", System.currentTimeMillis())).queryInterface(IRun.class);
                        newRun.status("Passed");
                        newRun.post();
                        currentRun = newRun;
                    }
                }
            }
        }

    }

    public static void updateSteps(String status, String description, String expected, String actual) {
        ++stepCount;
        IBaseFactory stepFactory = (IBaseFactory)currentRun.stepFactory().queryInterface(IBaseFactory.class);
        IStep currentStep = (IStep)stepFactory.addItem("Step " + stepCount).queryInterface(IStep.class);
        currentStep.status(status);
        currentStep.field("ST_DESCRIPTION", description);
        currentStep.field("ST_EXPECTED", expected);
        currentStep.field("ST_ACTUAL", actual);
        currentStep.post();
        if (status.equalsIgnoreCase("Failed")) {
            currentRun.status("Failed");
            currentRun.post();
            currentRun.refresh();
        }

    }

    public static void updateStepsWithScreenshot(String status, String description, String expected, String actual) {
        ++stepCount;
        IBaseFactory stepFactory = (IBaseFactory)currentRun.stepFactory().queryInterface(IBaseFactory.class);
        IStep currentStep = (IStep)stepFactory.addItem("Step " + stepCount).queryInterface(IStep.class);
        currentStep.status(status);
        currentStep.field("ST_DESCRIPTION", description);
        currentStep.field("ST_EXPECTED", expected);
        currentStep.field("ST_ACTUAL", actual);
        File file = FnLib.takeScreenshot(stepCount);
        IAttachmentFactory attachmentFactory = (IAttachmentFactory)currentStep.attachments().queryInterface(IAttachmentFactory.class);
        IAttachment screenshotToAttach = (IAttachment)attachmentFactory.addItem(file.getName()).queryInterface(IAttachment.class);
        IExtendedStorage extAttach = (IExtendedStorage)screenshotToAttach.attachmentStorage().queryInterface(IExtendedStorage.class);
        extAttach.clientPath(file.getParent());
        extAttach.save(file.getName(), true);
        screenshotToAttach.description(actual);
        screenshotToAttach.post();
        currentStep.post();
        if (status.equalsIgnoreCase("Failed")) {
            currentRun.status("Failed");
            currentRun.post();
            currentRun.refresh();
        }

    }
}
