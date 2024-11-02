package FinalProject;
import org.testng.annotations.Test;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;

public class DataDrivenTestingExcel {
	
  WebDriver driver;
  XSSFWorkbook workbook = new XSSFWorkbook();
  ExtentSparkReporter sparkReporter;
  ExtentReports extent;
  ExtentTest test; 
  
 
  @BeforeTest
  public void beforeTest() {
	  
	//WebDriver Setup
	  driver = new ChromeDriver();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT::::");
	  driver.manage().window().maximize();
	  driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10)); 
	  
	//ExtentReportSetup	   
	  sparkReporter = new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
	  sparkReporter.config().setDocumentTitle("Automation Report");
	  sparkReporter.config().setReportName("Test Execution Report");
	  sparkReporter.config().setTheme(Theme.STANDARD);
	  sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
	  extent = new ExtentReports();
	  extent.attachReporter(sparkReporter); 	
  }
  
  @Test
  public void AccessWebsiteAndRetreiveDatatoExcel() throws InterruptedException, IOException{
	//Navigate to tab "90 Days Bond"  
	WebElement Last90daysBond = driver.findElement(By.xpath("//*[@id='OBLIGATIONS_tab']"));
	Last90daysBond.click();

	//Identify the TableElement 
	List<WebElement> AllTablesElement = driver.findElements(By.xpath("//*[@id=\"report_OBLIGATIONS\"]/div/div[1]/table/tbody"));
	System.out.println("No of Tables in this Page is " +AllTablesElement.size());


    // Iterate through the first 5 tables sections
    for (int i = 1; i <=5  ; i++) {
    	WebElement Table = AllTablesElement.get(i);
        System.out.println(" Processing Table: " + (i));
        String methodName = new Exception().getStackTrace()[0].getMethodName();
        test= extent.createTest(methodName, "Links in the Table" +i);
        test.log(Status.INFO, "Table: " +i +" is Validated");
        test.addScreenCaptureFromPath(captureScreenshot("Clicked_Link_" +i));
        
     // Find all rows within the current <tbody>
        	List<WebElement> rows = Table.findElements(By.tagName("tr"));
        	System.out.println(" No of Rows in this section is " +rows.size());
        
        // Iterate through each row to find links and Click it 
        	for (WebElement row : rows) {
        		List<WebElement> links = row.findElements(By.tagName("a"));
        		for (WebElement link : links) {
        			String linkText = link.getText();
        			System.out.println("Clicking on Link: " + linkText);
                    link.click();
                    Thread.sleep(4000);
                    
                    test.log(Status.INFO, "Clicked Link: " +linkText );
                    test.addScreenCaptureFromPath(captureScreenshot("Clicked_Link_" + linkText));
                    Thread.sleep(4000);
                    
                 // Create a Sheet in the Workbook
                    XSSFSheet sheet = workbook.createSheet(linkText);
                    
                 //Switch to Frame 
                   driver.switchTo().frame(0);
                   WebElement WebPageExcelWebElement = driver.findElement(By.xpath("//div[2]/div[2]/table[2]/tbody"));
                   List<WebElement> WebPageExcelRows = WebPageExcelWebElement.findElements(By.tagName("tr"));
                    
                 // Iterate through each row of the data table
                   int ExcelrowIndex = 0;
                    for (WebElement WebPageExcelRow : WebPageExcelRows) {
                    Row ExcelsheetRow = sheet.createRow(ExcelrowIndex++);
              		List<WebElement> WebPageExcelColumns = WebPageExcelRow.findElements(By.tagName("td"));
              		  
              	     //iterate through columns
              		 int ExcelcolumnIndex = 0;
              		 for (WebElement WebPageExcelColumn : WebPageExcelColumns) { 
              			 String WebPageExcelCellValue = WebPageExcelColumn.getText();
              			 ExcelsheetRow.createCell(ExcelcolumnIndex++).setCellValue(WebPageExcelCellValue);
               			 System.out.println(WebPageExcelCellValue + "\t" );
              				  }	
              		 	System.out.println();
              		 }
                    
                    driver.switchTo().defaultContent();
                     WebElement CloseButton = driver.findElement(By.xpath("//button[@class='ui-button ui-corner-all ui-widget ui-button-icon-only ui-dialog-titlebar-close']"));
                     CloseButton.click();                               
                }        		
        		 try
                 {
                 FileOutputStream outputStream = new FileOutputStream("outputFile.xlsx");
                 test.log(Status.PASS, "Data is Written to Excel");
                 
                 //Close Stream
                 workbook.write(outputStream);   
                 } catch (IOException e) {
                     System.out.println("Error writing to Excel file: " + e.getMessage());
                     workbook.close();  
                 }  
        		}    
        	 }
   		}      
 
  //Screenshot Setup Section
  public String captureScreenshot(String screenshotName) throws IOException{
	  String FileSeparator = System.getProperty("file.separator");
	  String Extent_report_path = "."+FileSeparator+"Reports";
	  
	  File Src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	  String Screenshotname = "Screenshot"+Math.random()+".png";
	  File Dst = new File(Extent_report_path+FileSeparator+"Screenshots"+FileSeparator+Screenshotname);
	  
	  FileUtils.copyFile(Src, Dst);
	  
	  String absPath = Dst.getAbsolutePath();
	  System.out.println(" Screenshot and Report path is: "+absPath);
	  return absPath;
  }

  @AfterTest
  public void afterTest() {
	  extent.flush();
	  driver.close();
  }

}
