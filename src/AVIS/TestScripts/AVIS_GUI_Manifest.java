package AVIS.TestScripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;

import AVIS.CommonFunctions.*;

/**
'#################################################################################################################################
'## SCRIPT NAME:                           Avis_GUI_Manifest
'## BRAND:                                         AVIS
'## DESCRIPTION:                           Perform a Manifest of different Types to capture the report details in different regions.
'## FUNCTIONAL AREA :                      Manifest Screen
'## PRECONDITION:                          All the required Test Data should be available in Test Data Sheet.
'## OUTPUT:                                Reservation should be created successfully.

'##################################################################################################################################
 **/

public class AVIS_GUI_Manifest
{

	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport()
	{
		extent = Extentmanager.GetExtent();
		//test = extent.createTest("GUI");
	}

//	public static void main(String[] args) throws Throwable {
	@Test
	public void test() throws Exception
	{
		try
		{
			Properties prop = new Properties();
			FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\git\\ABG_GUI\\ABG_GUI_Automation\\src\\AVIS\\TestData\\TestDataABGGUI.properties");
			prop.load(fis);
			//WebDriver driver;
			ChromeDriver driver = new ChromeDriver();
			GUIFunctions functions = new GUIFunctions(driver);
			System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
			driver.navigate().to(prop.getProperty("AvisURL"));
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			Thread.sleep(2000);
			functions.txt_userid.sendKeys(prop.getProperty("USERID"));
			Thread.sleep(500);
			functions.txt_password.sendKeys(prop.getProperty("PASSWORD"));
			Thread.sleep(500);
			functions.btn_login.click();
		   for (int k=1; k<=20; k++)
		   {
			   Avis_GUI_Create_Reservation avis = new Avis_GUI_Create_Reservation();
				
				ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Avis\\AVIS_GUITestData_Manifest_Regression.xlsx");
				String Execute 	= rwe.getCellData("Manifest_Avis", k,  2);
				
				//********Delete the files in the folder********//
				File file = new File(prop.getProperty("ScreenshotAvis"));  

				String[] myFiles;    
				if (file.isDirectory()) {
				    myFiles = file.list();
				    for (int i = 0; i < myFiles.length; i++) {
				        File myFile = new File(file, myFiles[i]); 
				        myFile.delete();
				    }
				    int a = 28;
			   
			   String TCName 			= rwe.getCellData("Manifest_Avis", k,  4);
		       String clientURL 			= rwe.getCellData("Manifest_Avis", k,  6);
			   String USERID 					= rwe.getCellData("Manifest_Avis", k,  7);
			   String PASSWORD 				= rwe.getCellData("Manifest_Avis", k,  8);
			   String outSTA				= rwe.getCellData("Manifest_Avis", k,  9);
			   String thinClient = clientURL + outSTA;
//			   String COUNTER_NUM				= rwe.getCellData("Manifest_Avis", k, 13);
			   
			if (Execute.equals("N"))	
				driver.close();
			
			if (Execute.equals("Y"))
			{
				
				
				String ScreenshotPath = prop.getProperty("ScreenshotAvis");	
				
				/* Open GUI URL's */
				// System.out.println(" token URL value : " + tokenURL);
				//*******Screenshot path and test name*********//
				
				String testcasename = TCName;
				String xfilepath = prop.getProperty("ExcelPathAvis") +testcasename+ ".xlsx";
				test = extent.createTest(TCName);
				functions.openURL(thinClient);
				//driver.get(THINCLIENTURL + OUTSTATION);
				
				//functions.login(USERID, PASSWORD);

				driver.navigate().refresh();
				Thread.sleep(10000);
				driver.manage().window().maximize();
					
				driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
				Thread.sleep(2000);
				functions.ScreenCapturedate(ScreenshotPath,TCName);
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
				Thread.sleep(2000);
					
				String OneDay_LOR = rwe.getCellData("Manifest_Avis", k, 11);
				driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(OneDay_LOR);
				Thread.sleep(2000);
							
				driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
				Thread.sleep(8000);
				
				String DetailType = rwe.getCellData("Manifest_Avis", k, 10);
				System.out.println(DetailType);
				
			    driver.findElement(By.xpath("//*[@id=\"manifest_submenu\"]/div/div[3]/button")).click();;
			    Thread.sleep(3000);
			    System.out.println(" Manifest Type is " + DetailType);
				driver.findElement(By.linkText(DetailType)).click();
				Thread.sleep(3000);
				functions.ScreenCapturedate(ScreenshotPath,TCName);
				WebElement NoRes = driver.findElement(By.xpath("//*[@id='manifestGrid']/div[1]/div[2]/div/span"));
				NoRes.getText();
				String NoResFound = NoRes.getText();
				System.out.println(NoResFound);
				String ResVal	=	"No reservations found.";
				String NoData	=	"09-DATA NOT AVAILABLE FOR DTE ENTERED";
				String Norental = "No rentals found.";
				functions.ScreenCapturedate(ScreenshotPath,TCName);
	// Logout the application when the "No reservations found"
				if(NoResFound.equals(ResVal) || NoResFound.equals(NoData) || NoResFound.equals(Norental))
				{
					Date d= new Date();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//FileUtils.copyFile(src,new File( ScreenShotPath +TESTCASENAME+ sdf.format(d)+"_1Day"+".pgn"));
					continue;
					//functions.logout();
					//driver.close();
				}
				else
				{
					WebElement TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					String TotItems1= TotItems.getText();
					String str = TotItems1;
				    System.out.println(str);
				 
				    // remove the last character 
				    str = str.substring(9, str.length()-5);
				    rwe.setCellData("Manifest_Avis", k, 14, str);
					Thread.sleep(2000);
			        		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					String ResNo1= ResNo.getText();
						
	//		Thread.sleep(2000);
					rwe.setCellData("Manifest_Avis", k, 15, ResNo1);
					
					WebElement AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String AWD1= AWD.getText();
					System.out.println(AWD.getText());	
					rwe.setCellData("Manifest_Avis", k, 17, AWD1);
					
					WebElement Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					String Wizard1 = Wizard.getText();
					rwe.setCellData("Manifest_Avis", k, 18, Wizard1);
					
					WebElement Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					String Coupon1 = Coupon.getText();
					rwe.setCellData("Manifest_Avis", k, 46, Coupon1);
					
					WebElement PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					String PickUpLoc1 = PickUpLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 19, PickUpLoc1);
					
					WebElement PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					PickupDate.getText();
					String PickupDate1 = PickupDate.getText();
					System.out.println(PickupDate.getText());	
					rwe.setCellData("Manifest_Avis", k, 20, PickupDate1);
					
					WebElement ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					ReturnLoc.getText();
					String ReturnLoc1 = ReturnLoc.getText();
					System.out.println(ReturnLoc.getText());	
					rwe.setCellData("Manifest_Avis", k, 21, ReturnLoc1);
					
					WebElement ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					ReturnDateTime.getText();
					String ReturnDateTime1 = ReturnDateTime.getText();
					System.out.println(ReturnDateTime.getText());	
					rwe.setCellData("Manifest_Avis", k, 22, ReturnDateTime1);
					
					WebElement CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					CarClass.getText();
					String CarClass1 = CarClass.getText();
					System.out.println(CarClass.getText());	
					rwe.setCellData("Manifest_Avis", k, 23, CarClass1);
					
					WebElement BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					BookingDt.getText();
					String BookingDt1 = BookingDt.getText();
					System.out.println(BookingDt.getText());	
					rwe.setCellData("Manifest_Avis", k, 24, BookingDt1);
					Thread.sleep(2000);
					
	// Capture the screen shot of the report
					Date d1= new Date();
					SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src1= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//FileUtils.copyFile(src1,new File( ScreenShotPath +TESTCASENAME+ sdf1.format(d1)+"_1Day"+".pgn"));
					
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
	
	// Max days (2 days)of Manifest
					driver.findElement(By.xpath("//*[@id=\"menubar\"]/ul/li[5]/a")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).clear();
					Thread.sleep(2000);
					
					String TwoDay_LOR = rwe.getCellData("Manifest_Avis", k, 12);
					driver.findElement(By.xpath("//*[@id=\"manifestToDate\"]")).sendKeys(TwoDay_LOR);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//*[@id=\"getManifest\"]")).click();
					Thread.sleep(15000);
	
	//Capture the total number of items are displayed		
					WebElement max_TotItems = driver.findElement(By.xpath("//*[@id=\"manifestGrid\"]/div[2]/div[2]/div/span"));
					String max_TotItems1= max_TotItems.getText();
										
					String max_str = max_TotItems1;
					max_str = max_str.substring(9, max_str.length()-5);
					rwe.setCellData("Manifest_Avis", k, 25, max_str);
	
	//Select a one of the record to open in report format		
					driver.findElement(By.xpath("//div[@class='ui-grid-canvas']/div[1]/div[1]/div[1]/div[1]")).click();
					driver.findElement(By.xpath("//*[@id=\"menulist:manifestlink\"]/div[2]/div/div/div[2]/button")).click();
					Thread.sleep(3000);
					
					WebElement max_ResNo = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[1]/div/span"));
					String max_ResNo1= max_ResNo.getText();
					rwe.setCellData("Manifest_Avis", k, 26, max_ResNo1);
					
					WebElement max_AWD = driver.findElement(By.xpath("//*[@id='resDetailPanel']/div[2]/div[2]/div[1]/div/div[3]/div[2]"));
					String max_AWD1= max_AWD.getText();
					rwe.setCellData("Manifest_Avis", k, 28, max_AWD1);
					
					WebElement max_Wizard = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[1]/div/div[5]/div[2]"));
					String max_Wizard1 = max_Wizard.getText();
					rwe.setCellData("Manifest_Avis", k, 29, max_Wizard1);
					
					WebElement max_Coupon = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[2]/div[1]/div/div[7]/div[2]"));
					max_Coupon.getText();
					
					WebElement max_PickUpLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[3]/div[2]/span"));
					String max_PickUpLoc1 = max_PickUpLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 30, max_PickUpLoc1);
					
					WebElement max_PickupDate = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[3]/div/div[5]/div[2]/span"));
					String max_PickupDate1 = max_PickupDate.getText();
					rwe.setCellData("Manifest_Avis", k, 31, max_PickupDate1);
					
					WebElement max_ReturnLoc = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[3]/div[2]/span"));
					String max_ReturnLoc1 = max_ReturnLoc.getText();
					rwe.setCellData("Manifest_Avis", k, 32, max_ReturnLoc1);
					
					WebElement max_ReturnDateTime = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[3]/div[5]/div/div[5]/div[2]/span"));
					String max_ReturnDateTime1 = max_ReturnDateTime.getText();
					rwe.setCellData("Manifest_Avis", k, 33, max_ReturnDateTime1);
					
					WebElement max_CarClass = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[4]/span"));
					String max_CarClass1 = max_CarClass.getText();
					rwe.setCellData("Manifest_Avis", k, 34, max_CarClass1);
					
					WebElement max_BookingDt = driver.findElement(By.xpath("//*[@id=\"resDetailPanel\"]/div[2]/div[1]/div[2]/div/div[16]/span"));
					String max_BookingDt1 = max_BookingDt.getText();
					rwe.setCellData("Manifest_Avis", k, 35, max_BookingDt1);
					Thread.sleep(2000);
					
	// Capture the screen shot of the report
					Date d2= new Date();
					SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
					File src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//FileUtils.copyFile(src2,new File( ScreenShotPath +TESTCASENAME+ sdf2.format(d2)+"_2Day"+".pgn"));
					//Select a row from the displayed open reservations
					driver.findElement(By.id("mdetailCloseButton")).click();
					Thread.sleep(2000);
					
	// Logout the application
					test = extent.createTest(TCName);
					
					if (rwe.getCellData("Manifest_Avis", k, 22).isEmpty())
					{
						rwe.setCellData("Manifest_Avis", k, 23, "FAIL");
						test.log(Status.FAIL, "Fail");
					}
					else
					{
						test.log(Status.PASS, "Pass");
						rwe.setCellData("Manifest_Avis", k, 23, "PASS");
					}

					
					//Capturing all the screenshots in excel sheet
					FileOutputStream fileOut = null;
					int cntr =0;
					int row = 0;
					try {

					       Workbook wb = new XSSFWorkbook();
					      Sheet sheet = wb.createSheet("Ouput");
					       // FileInputStream obtains input bytes from the image file
					       String[] pathnames;

					       // Creates a new File instance by converting the given pathname string
					       // into an abstract pathname
					       File f = new File(prop.getProperty("ScreenshotAvis"));

					       // Populates the array with names of files and directories
					       pathnames = f.list();

					       // For each pathname in the pathnames array
					       for (String pathname : pathnames) {
					             // Print the names of files and directories

					             InputStream inputStream = new FileInputStream(
					            		 prop.getProperty("ScreenshotAvis")+pathname);
					             byte[] bytes = IOUtils.toByteArray(inputStream);
					             
					             // Adds a picture to the workbook
					       
					             int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
					             
					             // close the input stream
					             
					             // Returns an object that handles instantiating concrete classes
					             CreationHelper helper = wb.getCreationHelper();
					             // Creates the top-level drawing patriarch.
					             Drawing drawing = sheet.createDrawingPatriarch();

					             // Create an anchor that is attached to the worksheet
					             ClientAnchor anchor = helper.createClientAnchor();

					             // create an anchor with upper left cell _and_ bottom right cell
					             System.out.println(cntr);
					             anchor.setCol1(cntr); // Column B
					             anchor.setRow1(cntr+row ); // Row 3
					           

					             // Creates a picture
					             Picture pict = drawing.createPicture(anchor, pictureIdx);

					             // Reset the image to the original size
					             pict.resize(); //don't do that. Let the anchor resize the image!

					             // Create the Cell B
					             Cell cell = sheet.createRow(cntr).createCell(cntr);

					             // Write the Excel file

					             
					             fileOut = new FileOutputStream(xfilepath);
					             wb.write(fileOut);
					             row = row +40 ;
					             
					             inputStream.close();
					       }
					       //fileOut.close();

					} catch (IOException ioex) {
					       System.out.println(ioex);
					}
				
					
				}
			}//End of If statement
			
			else
			{
			   System.out.println("Execution status is N for iteration "+k+"...");
		    }
		}
	}//End of for statement
		   functions.logout();
			Thread.sleep(1000);
			functions.closeWindows();
		}
	finally
	{
		// TODO: handle finally clause
		extent.flush();
	}
  }
}