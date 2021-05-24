package AVIS.TestScripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

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
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;

import AVIS.CommonFunctions.*;
import AVIS.CommonFunctions.ReadWriteExcel;
import Payless.TestScripts.Payless_GUI_Prepaid_Voucher_Rentals;

/**
 * '#############################################################################################################################
 * '## SCRIPT NAME: AVIS_GUI_Extended_Rentals '## BRAND: AVIS '## DESCRIPTION:
 * Creating RA numbers for Extended Rentals '## FUNCTIONAL AREA :
 * Checkout,Checkin, Display Rental screen '## PRECONDITION:All the required
 * Test Data should be available in Test Data Sheet. '##OUTPUT: Extended Rentals
 * should be created successfully.
 * 
 * HISTORY 12-SEP-2018 - GUIFunctions class created for GUI Common
 * functionalities and CR functionality
 * '#############################################################################################################################
 **/

public class AVIS_GUI_Extended_Rentals {

	public void clickRateshopSearchBtn(ChromeDriver driver) {
		JavascriptExecutor jse = (JavascriptExecutor) driver;
		String clickSearchJS = "document.getElementById('searchCommandLinkResRateCode').click()";
		jse.executeScript(clickSearchJS);
	}

	ExtentReports extent;
	ExtentTest test;

	@BeforeTest
	public void startReport() {

		extent = Extentmanager.GetExtent();
		// test = extent.createTest("GUI");

	}

	@SuppressWarnings("unlikely-arg-type")
	// public static void main(String[] args) throws IOException, Exception,
	// FileNotFoundException {

	@Test
	public void test() throws Exception {
		
			try{
				Properties prop = new Properties();
				FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\git\\ABG_Sele_2020\\ABG_GUI_Automation\\src\\AVIS\\TestData\\TestDataABGGUI.properties");
				prop.load(fis);
				//WebDriver driver;
				
				ChromeOptions chromeOptions= new ChromeOptions(); 
				chromeOptions.setBinary("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"); 
				System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
				ChromeDriver driver = new ChromeDriver(chromeOptions);
				GUIFunctions functions = new GUIFunctions(driver);
		 
				
				driver.navigate().to(prop.getProperty("AvisURL"));
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				Thread.sleep(2000);
				functions.txt_userid.sendKeys(prop.getProperty("USERID"));
				Thread.sleep(500);
				functions.txt_password.sendKeys(prop.getProperty("PASSWORD"));
				Thread.sleep(500);
				functions.btn_login.click();
				
					/* Login */
				Thread.sleep(3000);
			// Read input from excel
			///int intRowCount     = 100;
				for (int k = 1; k <= 100; k++)
				{
					AVIS_GUI_Extended_Rentals avis = new AVIS_GUI_Extended_Rentals();
					ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Avis\\AVIS_GUITestData_Extended_Rental.xlsx");
														
					String Execute = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 2);
					
					//********Delete the files in the folder********//
					File file = new File(prop.getProperty("ScreenshotAvis"));  

					String[] myFiles;    
					if (file.isDirectory()) {
					    myFiles = file.list();
					    for (int i = 0; i < myFiles.length; i++) {
					        File myFile = new File(file, myFiles[i]); 
					        myFile.delete();
					    }
					}
					
					int a = 27;
					
					// ChromeDriver driver = new ChromeDriver();
					// GUIWebDriverFunctions wdfunctions = new GUIWebDriverFunctions(driver);
					driver.manage().window().maximize();
					driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
					System.out.println(" iteration " + k);
					String TCName = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 4);
					String tokenURL = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 6);
					String clientURL = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 7);
					String outSTA = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 8);
					String thinClient = clientURL + outSTA;
					String uName = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 9);
					String pswd = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 10);
					String Resno = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 11);
					String Coemail = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 12);
					String COMILEAGE1 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 13);
					String COMVA = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 14);
					String RANumber_1 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 15);
					String COCountry 		= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 16);
					String COState			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 17);
					String CODRLICNO 		= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 18);
					String CODOB			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 19);
					String COCOMPANY		= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 20);
					String COADDR1			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 21);
					String COADDR2			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 22);
					String COADDR3			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 23);
					String COCONTACTINFO   = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 24);
					String COEMAIL			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 25);
					String COMOP			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 26);
					String COCCDC			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 27);
					String CARDNO			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 28);
					String COMM			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 29);
					String COYY			= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 30);
					String MOPREASON		= rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 31);
					String RANumber_2 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 32);
					String MilageIn2 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 33);
					String MilageIn3 = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 34);
					String Action_code_CI = rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 35);
					
				if (Execute.equals("Y")) {
					

					String ScreenshotPath = prop.getProperty("ScreenshotAvis");	
					
					//* Open GUI URL's */
						// System.out.println(" token URL value : " + tokenURL);
						//*******Screenshot path and test name*********//
						
						String testcasename = TCName;
						String xfilepath = prop.getProperty("ExcelPathRentalAvis") +testcasename+ ".xlsx";
						test = extent.createTest(TCName);
							
						functions.openURL(thinClient);

					/* Open GUI URL's */
					/*driver.get(thinClient);

					/* Login */
					//functions.login(uName, pswd);
					driver.navigate().refresh();
					functions.navigateToTab("CheckOut");
					Thread.sleep(8000L);

					// Step1: On checkout screen--> click delay toggle and enter Res no in the field
					// and search
					driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[1]/div[1]/div[1]/button[1]/span[1]")).click();
					Thread.sleep(4000);
					driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[1]/div[1]/div[1]/div[1]/ul[1]/li[2]/span[2]")).click();
					
					//driver.findElement(By.xpath("//span[@id='checkoutToggle']")).isDisplayed();//gui 2 link
					//driver.findElement(By.xpath("//span[@id='checkoutToggle']")).click();//gui 2 link
					//driver.findElement(By.id("delayBtn")).isDisplayed();//gui 1 link
					//driver.findElement(By.id("delayBtn")).click();//gui 1 link
					Thread.sleep(7000);
					
					// Enter the res no in search field and click on search button
					driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).click();
					driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).sendKeys(Resno);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@ng-click='directSearch()']")).click();
					Thread.sleep(15000);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).click();
					  String Country = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).getAttribute("value");
					  System.out.println("The country is " + Country);
					  if (Country.isEmpty())
					  {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseCountry']")).sendKeys(COCountry);
					 System.out.println("Entered the Country");
					  }
					  
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).click();
					 String State =driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).getAttribute("value");
					 System.out.println("The state is " +State );
					 if (State.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseState']")).sendKeys(COState);
					 System.out.println("Entered the State");
					 }
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).click();
					 String Licence = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).getAttribute("value");
					 System.out.println("The Licence entered id " +Licence);
					 if (Licence.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:licenseNumber']")).sendKeys(CODRLICNO);
					 System.out.println("Entered the Licence Number");
					 }
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).click();
					 String DOB= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).getAttribute("value");
					 System.out.println("The Date od Birth entered is "+DOB);
					 if (DOB.isEmpty())
					 {	 
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']")).sendKeys(CODOB);
					 System.out.println("Entered the Date of Birth");
					 }
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).click();
					 String Company = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).getAttribute("value");
					 System.out.println("The Company entered is "+Company);
					 if (Company.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:company']")).sendKeys(COCOMPANY);
					 System.out.println("Entered the Company");
					 }
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).click();
					 String Add1 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).getAttribute("value");
					 if (Add1.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address1']")).sendKeys(COADDR1);
					 System.out.println("Entered the Address 1");
					 }
					 String Add2= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).getAttribute("value");
					 if (Add2.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address2']")).sendKeys(COADDR2);
					 System.out.println("Entered the Address 2");
					 }
					 String Add3 = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).getAttribute("value");
					 if (Add3.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:address3']")).sendKeys(COADDR3);
					 System.out.println("Entered the Address 3");
					 }
					 String Contact =driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).getAttribute("value");
					 if (Contact.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:contactInfo']")).sendKeys(COCONTACTINFO);
					 System.out.println("Entered the Contact Info");
					 }
					 if(driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).isEnabled())
					 {
					 String Email= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).getAttribute("value");
					 if (Email.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:wizconEmailInput']")).sendKeys(COEMAIL);
					 System.out.println("Entered the Email");
					 }
					 }
					
					 
					// driver.findElement(By.xpath("//input[@id='glyphicon glyphicon-plus']")).click();
					 /*String prepay = driver.findElement(By.xpath("//*[@id='menulist:checkoutContainer:checkoutForm:prePay:mopAmount']")).getAttribute("value");
					 System.out.println(prepay);
					 driver.findElement(By.xpath("//*[@id='menulist:checkoutContainer:checkoutForm:prePay:mopAmount']")).clear();
					 driver.findElement(By.xpath("//*[@id='menulist:checkoutContainer:checkoutForm:prePay:mopAmount']")).sendKeys(PREPAY);*/
					 
					 Thread.sleep(2000);
					 Select s= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:payMethod")));
					 s.selectByValue(COMOP);
					 System.out.println("Entered the Method of Payment");
					 Select s2= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccType")));
					 s2.selectByValue(COCCDC);
					 System.out.println("Entered the Type of Payment");
					 String CardNumber = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).getAttribute("value");
					 if (CardNumber.isEmpty())
					 {
						
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccNumber']")).sendKeys(CARDNO);
					 System.out.println("Entered the Card Number");
					 }
					 String Month = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).getAttribute("value");
					 if (Month.isEmpty())
					 {
					 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireMonth']")).sendKeys(COMM);
						System.out.println("Entered the Expiry Month");
					 }
					 String year= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).getAttribute("value");
					 if (year.isEmpty())
					 {
						driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:swipe:ccExpireYear']")).sendKeys(COYY);
						System.out.println("Entered the Expiry Year");
					 }
					 	Thread.sleep(4000);
						Select s3= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason")));
						
						s3.selectByVisibleText(MOPREASON);
						
						System.out.println("Entered the Reason for payment");
						//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:mopCCI']")).sendKeys("1234");
						//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:authorization']")).sendKeys("OK/1234");
						//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).clear();
						//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).sendKeys("GB/C");
							
								
								
						String RateCode = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).getAttribute("value");
						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 63,RateCode);
						driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(COMVA);
						System.out.println("Entered the MVA number");
						driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
						Thread.sleep(2000);
						driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
						Thread.sleep(2000);
						driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(MilageIn2);
						System.out.println("Entered the Mileage");
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						if(driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).isDisplayed())
						{
						driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleButton']")).click();
						}
						Thread.sleep(7000);
						
						/*driver.findElement(By.xpath("//*[@id='checkoutDialog_validatePrepaymentCreditCardForm:cardnumber']")).click();
						driver.findElement(By.xpath("//*[@id='checkoutDialog_validatePrepaymentCreditCardForm:cardnumber']")).sendKeys("6020");
						Thread.sleep(3000);
						driver.findElement(By.xpath("//*[@id='checkoutDialog_validatePrepaymentCreditCardForm:validatePrepaymentCreditCardContinueButton']")).click();*/
						
						if(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).isDisplayed())
						{
						String OptMSG= driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getAttribute("value");
						System.out.println("The Output message in vechicle continue screen is "+ OptMSG);
						if(OptMSG.equals("NET RATE ERROR")) 
						{
						System.out.println("ERROR MESSAGE "+OptMSG+" is Displayed");
						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 70, OptMSG);
						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 71, "FAIL");
						test.log(Status.FAIL, "Fail");
						
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						System.out.println("The Screenshot is taken");
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton")).click();
						driver.close();
						break;
						}
						else
							if (OptMSG.equals("START DATE INVALID FOR RATE"))
							{
							System.out.println("ERROR MESSAGE "+OptMSG+" is Displayed");
							rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 70, OptMSG);
							rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 71, "FAIL");
							test.log(Status.FAIL, "Fail");
							
							functions.ScreenCapturedate(ScreenshotPath,TCName);
							System.out.println("The Screenshot is taken");
							driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton")).click();
							driver.close();
							break;	
								
							}
							else
								if(OptMSG.equals("INVALID DATA - REENTER"))
								{
									driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePRP")).sendKeys("V10000");
									
									driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
									Thread.sleep(10000);
								}
						}
						
						
							//Thread.sleep(20000);
						//functions.ScreenCapturedate(ScreenshotPath,TCName);
						if (driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).isDisplayed())
						{
	 				  	driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
	 				  	System.out.println("Clicked on Continue Button");
						Thread.sleep(20000);
						}
						
						Thread.sleep(20000);
						
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						if(driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).isDisplayed())
						{
						driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).click();
						System.out.println("Clicked on Complete Checkout Button");
						Thread.sleep(8000);
						}
						Thread.sleep(5000);
					
					// Step2: Enter the soft prompt(email address)and click on continue
					String email = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut"))
							.getText();
					if (email.contains("PLEASE INPUT CUSTOMER EMAIL ADDRESS")) {
						// boolean emailfield =
						// driver.findElement(By.xpath("//input[@class='form-control ng-pristine
						// ng-valid allowBackspace ng-touched']")).isDisplayed();
						// System.out.println(emailfield);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:wizconEmail")).click();
						Thread.sleep(4000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:wizconEmail")).sendKeys(Coemail);
						Thread.sleep(2000);
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:wizconSubmitButton']"))
								.isDisplayed();
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:wizconSubmitButton']"))
								.click();
						Thread.sleep(15000);
					}

					// If DOB is not entered
					/*String DOB = driver
							.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']"))
							.getAttribute("value");
					System.out.println("The Date od Birth entered is " + DOB);
					if (DOB.isEmpty()) {
						driver.findElement(
								By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:dateOfBirth']"))
								.sendKeys("06/09/77");
						System.out.println("Entered the Date of Birth");
					}*/
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Step3:Click on delay continue button on checkout screen
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).click();
					Thread.sleep(8000);

					// Step4:Click on Ok Button on the pop up window displayed
					/*String partialcheckout = driver.switchTo().alert().getText();
					if (partialcheckout.contains("Are you sure you want to do a partial checkout?")) {
						driver.switchTo().alert().accept();
						Thread.sleep(5000);
					}*/

					Thread.sleep(9000);
					// Enter MVA number
					String comilenmva = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut"))
							.getText();
					if (comilenmva.contains("ERROR - SEE HIGHLIGHTED FIELDS")) {
						// Enter all the required fields checkout miles and MVA number
						Thread.sleep(5000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.isDisplayed();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.click();
						Thread.sleep(5000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
								.sendKeys(COMILEAGE1);
						Thread.sleep(2000);
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.isDisplayed();
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.click();
						driver.findElement(By.xpath(
								"//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue']"))
								.sendKeys(COMVA);
						Thread.sleep(2000);
						// Click on Continue Button
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']"))
								.isDisplayed();
						driver.findElement(
								By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']"))
								.click();
						Thread.sleep(8000);
					}
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Enter the valid RA number in the field
					String RANumber1 = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
					if (RANumber1.contains("NON NUMERIC RA NUMBER")) {
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(RANumber_1);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
						Thread.sleep(5000);
					}

					String documentnoinuse = driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
					if (documentnoinuse.contains("DOCUMENT NUMBER ALREADY IN USE")) {
						while(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText().contains("C/O CAR INFO DISCREPANCIES") == false) {

							
							//String new_Ra2 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
							//int new_RA = get_RA_Number(Integer.parseInt(RANumber_1));
							int new_RA = get_RA_Number(Integer.parseInt(RANumber_1));
							
							driver.findElement(
										By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
								driver.findElement(
										By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue"))
										.sendKeys(Integer.toString(new_RA));
								driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton"))
										.click();
								Thread.sleep(5000);

							}
					
						} 
					functions.ScreenCapturedate(ScreenshotPath,TCName);

					// For message as C/O CAR INFO DISCREPANCIES
					String afterRa1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText();
					if (afterRa1.contains("C/O CAR INFO DISCREPANCIES")) or (afterRa1.contains("DO NOT RENT THIS CAR"));
					{
						/*driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(MilageIn2);
						Thread.sleep(2000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue")).clear();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue")).click();
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValue")).sendKeys(COMVA);*/
						Thread.sleep(2000);
						driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
						Thread.sleep(10000);
					}

					// Click on complete checkout button
					driver.findElement(By.id("footerForm:completeCheckoutButton")).isDisplayed();
					driver.findElement(By.id("footerForm:completeCheckoutButton")).click();
					Thread.sleep(10000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//Click on Rental charges
					Thread.sleep(5000);
					if(driver.findElement(By.id("rentalCharges:completeCheckoutBtn")).isDisplayed())
					{
						driver.findElement(By.id("rentalCharges:completeCheckoutBtn")).click();
						Thread.sleep(7000);
					}
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Capture back the values into Xls
					String strCOLNFNGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 36, strCOLNFNGetText);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
					String strCOVehicleModelGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 37, strCOVehicleModelGetText);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
					String strCOResNoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 38, strCOResNoGetText);
					System.out.println(" Reservation No value is " + strCOResNoGetText);
					String strCOMVAGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 39, strCOMVAGetText);
					System.out.println(" MVA No value is " + strCOMVAGetText);
					String strCORANoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 40, strCORANoGetText);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println(" 1st RA Number value is  " + strCORANoGetText);
					String strCOLicensePlateNoGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 41, strCOLicensePlateNoGetText);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
					String strCOEstimatedCostGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 42, strCOEstimatedCostGetText);
					System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
					String strCOSystemMsgGetText = driver.findElement(By.xpath(
							"//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 43, strCOSystemMsgGetText);
					System.out.println(" 1st CheckOut Complete System Message value is " + strCOSystemMsgGetText);
					Thread.sleep(7000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Click on View Rental Agreement button
					System.out.println(driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).isDisplayed());
					
					driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).click();
					Thread.sleep(15000);

					// Click on Extended button in display Rental screen
					//driver.findElement(By.xpath("//form[@id='menulist:dispormodSubMenu']/div[8]/button/span")).isDisplayed();
					//driver.findElement(By.xpath("//form[@id='menulist:dispormodSubMenu']/div[8]/button/span")).click();
					driver.findElement(By.xpath("//button[@class='rentalExtendClass']")).isDisplayed();
					driver.findElement(By.xpath("//button[@class='rentalExtendClass']")).click();
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Enter RAnumber2 and Mileagein in extended rental popup window
					String extendedrollover = driver.findElement(By.xpath("//div[@id='extendedrollover']/div/h1")).getText();
					if (extendedrollover.contains("Extended Rental")) {
						//int new_RA1 = get_RA_Number(Integer.parseInt(RANumber_2));
						driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).isDisplayed();
						//driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).click();
						Thread.sleep(2000);
						driver.findElement(By.xpath("//input[@class='form-control ng-pristine ng-valid ng-valid-maxlength allowBackspace ng-touched']")).sendKeys(RANumber_2);
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).isDisplayed();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).click();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedrollovermileage")).sendKeys(MilageIn3);
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedDone")).isDisplayed();
						driver.findElement(By.id("minileaserental:minileaserentalPopup:extendedDone")).click();
						Thread.sleep(9000);
						//Click on the "Ok" button "ERROR - SEE HIGHLIGHTED FIELDS : MILEAGE"
						driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).isDisplayed();
						driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).click();
						Thread.sleep(15000);
						functions.ScreenCapturedate(ScreenshotPath,TCName);
					}
					
					//Click on delayed toggle
					/*driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					//Select purchase fuel
					Select purchasefuel = new Select(driver.findElement(By.xpath("//select[@onchange='disableFuelLevel(this)']")));
					purchasefuel.selectByVisibleText("Yes");
					Thread.sleep(3000);
					
					//click on "complete checkin " button 
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).click();
					Thread.sleep(10000);*/
					Thread.sleep(15000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					//Enter the Second RA number and click on Roll over button
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendednewra']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendednewra']")).click();
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendednewra']")).sendKeys(RANumber_2);
					Thread.sleep(3000);
					//Enter the milage and click on done button
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendedrollovermileage']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendedrollovermileage']")).click();
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendedrollovermileage']")).sendKeys(MilageIn3);
					//Click on Roll over button
					Thread.sleep(7000);
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendedDone']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:extendedDone']")).click();
					Thread.sleep(9000);
					
					//Click on the "Ok" button "ERROR - SEE HIGHLIGHTED FIELDS : MILEAGE"
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='minileaserental:minileaserentalPopup:mleaseCheckInExceptionOk']")).click();
					Thread.sleep(15000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					
					//clcik on Complete checkin button
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).click();
					Thread.sleep(15000);
					
					//Enter 2nd RA number in Extended rental screen and click on Rollover button
					driver.findElement(By.xpath("//input[@id='checkinDialogs:checkInSuccessForm:nextraNumberExtnd']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinDialogs:checkInSuccessForm:nextraNumberExtnd']")).sendKeys(RANumber_2);
					Thread.sleep(5000);
					//click on rollover button
					driver.findElement(By.xpath("//button[@id='checkinDialogs:checkInSuccessForm:extndDlgRolloverButton']")).isDisplayed();
					driver.findElement(By.xpath("//button[@id='checkinDialogs:checkInSuccessForm:extndDlgRolloverButton']")).click();
					Thread.sleep(15000);
					
					//Click on Continue button
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
					Thread.sleep(15000);
					
					//click on Ok button
					driver.findElement(By.xpath("//input[@id='rolloverCheckoutExceptionOk']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='rolloverCheckoutExceptionOk']")).click();
					Thread.sleep(15000);
					
					//enter MVA and Mile
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(COMVA);
					System.out.println("Entered the MVA number");
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(MilageIn3);
					System.out.println("Entered the Mileage");
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					Thread.sleep(4000);
					////click on Ok button
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).click();
					Thread.sleep(15000);
					
					//Click on Continue button
					driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
					Thread.sleep(10000);
					
					//Click on Complete checkout button
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).click();
					Thread.sleep(10000);
					
					//Click on Rental charges Complete checkout button
					driver.findElement(By.xpath("//input[@id='rentalCharges:completeCheckoutBtn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='rentalCharges:completeCheckoutBtn']")).click();
					Thread.sleep(10000);
					
					// Capture back the values into Xls
					String strCOLNFNGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 44, strCOLNFNGetText1);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText1);
					String strCOVehicleModelGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 45, strCOVehicleModelGetText1);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText1);
					String strCOResNoGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 46, strCOResNoGetText1);
					System.out.println(" Reservation No value is " + strCOResNoGetText1);
					String strCOMVAGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 47, strCOMVAGetText1);
					System.out.println(" MVA No value is " + strCOMVAGetText1);
					String strCORANoGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 48, strCORANoGetText1);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println(" 2nd RA Number value is  " + strCORANoGetText1);
					String strCOLicensePlateNoGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 49, strCOLicensePlateNoGetText1);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText1);
					String strCOEstimatedCostGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 50, strCOEstimatedCostGetText1);
					System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText1);
					String strCOSystemMsgGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 51, strCOSystemMsgGetText1);
					System.out.println(" 1st CheckOut Complete System Message value is " + strCOSystemMsgGetText1);
					Thread.sleep(7000);
					functions.ScreenCapturedate(ScreenshotPath,TCName);
					// Click on View Rental Agreement button
					System.out.println(driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).isDisplayed());
					
					driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDisplayRentalButton']")).click();
					Thread.sleep(15000);

					
					/*// Capture back the values into Xls
					String strCOLNFNGetText1 = driver.findElement(By.xpath("//label[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteName']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 44, strCOLNFNGetText1);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText1);
					
					String strCOVehicleModelGetText1 = driver.findElement(By.xpath("//label[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteVehicle']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 45, strCOVehicleModelGetText1);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText1);
			
					String strCOMVAGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteMVA']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 46, strCOMVAGetText1);
					System.out.println(" MVA No value is " + strCOMVAGetText1);
					
					String strCORANoGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteRANumber']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 47, strCORANoGetText1);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println("2nd RA Number value is " + strCORANoGetText1);
					
					String strCOLicensePlateNoGetText1 = driver.findElement(By.xpath(
							"//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']"))
							.getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 48, strCOLicensePlateNoGetText1);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText1);
					
					String strCOEstimatedCostGetText1 = driver.findElement(By.xpath("//span[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteEstimatedCost']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 49, strCOEstimatedCostGetText1);
					System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText1);
					
					String strCOSystemMsgGetText1 = driver.findElement(By.xpath("//textarea[@id='checkinRolloverDialog:checkinRolloverCompleteForm:checkoutCompleteOut']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 50, strCOSystemMsgGetText1);
					System.out.println(" 2nd CheckOut Complete System Message value is " + strCOSystemMsgGetText1);
					Thread.sleep(7000);
					
					//Click on Done button
					driver.findElement(By.xpath("//input[@id='checkinRolloverDialog:checkinRolloverCompleteForm:mleasecheckoutCompleteDoneButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRolloverDialog:checkinRolloverCompleteForm:mleasecheckoutCompleteDoneButton']")).click();
					Thread.sleep(8000); */
					
					//Click on toggle button and Enter the RA_number2 in search field and click on search in checkin screen
					//driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					//driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).click();
					driver.findElement(By.xpath("//input[@id='checkinSearchString']")).sendKeys(RANumber_2);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='checkinSearchCommandLink']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinSearchCommandLink']")).click();
					Thread.sleep(10000);
					
					//Click on delayed toggle
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).isDisplayed();
					driver.findElement(By.xpath("//span[@id='checkinToggle']")).click();
					
					//Select purchase fuel
					Select purchasefuel1 = new Select(driver.findElement(By.xpath("//select[@onchange='disableFuelLevel(this)']")));
					purchasefuel1.selectByVisibleText("Yes");
					Thread.sleep(3000);
					
					
					//click on "complete checkin " button 
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='footerForm:completeCheckIn']")).click();
					Thread.sleep(10000);
					
					
					//Enter the mileage in pop up displayed and click on Continue button
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).click();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).clear();
					//driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptTable:0:repromptValue']")).sendKeys(MilageIn3);
					Thread.sleep(2000);
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
					Thread.sleep(9000);
					
					//Click on Continue button again in next window
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkinRepromptDialog:checkInRepromptForm:repromptSubmitButton']")).click();
					Thread.sleep(9000);
					
					//Capture the checkin details back to xls
					String strCOLNFNGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successName']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 51, strCOLNFNGetText2);
					System.out.println(" Last Name First Name value is " + strCOLNFNGetText2);
					String strCOVehicleModelGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successVehicle']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 52, strCOVehicleModelGetText2);
					System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText2);
					String strCORANoGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successRA']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 53, strCORANoGetText2);
					//rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 16, strCORANoGetText);
					System.out.println(" 2nd RA Number value is " + strCORANoGetText2);
					String strCOLicensePlateNoGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successPlate']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 54, strCOLicensePlateNoGetText2);
					System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText2);
					String strCOSystemMsgGetText2 = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkinDialogs:checkInSuccessForm:successOutMsg']")).getText();
					Thread.sleep(7000);
					rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 55, strCOSystemMsgGetText2);
					System.out.println(" 2nd Checkin Complete System Message value is " + strCOSystemMsgGetText2);
					Thread.sleep(7000);
					
					
					//Click on done button on the last screen or view rental
					//driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).isDisplayed();
					//driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:doneCompleteCheckIn']")).click();
					driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:viewCheckInRental']")).isDisplayed();
					driver.findElement(By.xpath("//input[@id='checkInSuccessDlg:checkInSuccessForm:viewCheckInRental']")).click();
					Thread.sleep(18000);
					
					// taking screenshot
					String ScreenShotPath = "C:\\Users\\cmn\\Desktop\\Selenium\\Screenshots\\Avis_GUI_Extended_Rentals\\";
					functions.ScreenCapturedate(ScreenShotPath, testcasename);
					

					/*
					 * Log out and close tabs
					 */

					functions.logout();
					Thread.sleep(1000);
					functions.closeWindows();
					test = extent.createTest(testcasename);
					if (rwe.getCellData("AVIS_GUI_Extended_Rentals", k, 38).isEmpty()) {

						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 57, "FAIL");
						test.log(Status.FAIL, "Fail");
					} else {
						rwe.setCellData("AVIS_GUI_Extended_Rentals", k, 57, "PASS");
						test.log(Status.PASS, "Pass");
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
					       fileOut.close();

					} catch (IOException ioex) {
					       System.out.println(ioex);
					}
			 
			}//end of if statement
				 else {
					System.out.println("Execution status is N for iteration " + k + "...");
				}

				}	

		} finally {
			extent.flush();
			// TODO: handle finally clause
		}
	}

	private void or(boolean contains) {
		// TODO Auto-generated method stub
		
	}

	public static int get_RA_Number(int RA_Num) {

		if (Integer.toString((RA_Num)).endsWith("6")) {
			RA_Num = RA_Num + 4;
			System.out.println(RA_Num);
		}

		else {
			RA_Num = RA_Num + 11;
			//System.out.println(RA_Num);
		}
		return RA_Num;
		
		
	}
}
