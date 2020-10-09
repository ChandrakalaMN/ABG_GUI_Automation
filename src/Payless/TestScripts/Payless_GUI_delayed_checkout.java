package Payless.TestScripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
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
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.gui.report.Extentmanager;
import com.ibm.icu.text.DateFormat;
import com.ibm.icu.text.SimpleDateFormat;

import AVIS.CommonFunctions.GUIFunctions;
import AVIS.CommonFunctions.ReadWriteExcel;

/* /**
 * '#############################################################################################################################
 * '## SCRIPT NAME: Payless_GUI_delayed_checkout '## BRAND: Payless '## DESCRIPTION:
 * Create delayed Rentals for reservations
 * products. '## FUNCTIONAL AREA : Reservation Rates Screen '## PRECONDITION:
 * All the required Test Data should be available in Test Data Sheet. '##
 * OUTPUT: Delayed rentals should be created successfully.
 * 
 * 
 * HISTORY 17-July-2020 - GUIFunctions class created for GUI Common
 * functionalities and CR functionality
 * '#############################################################################################################################
 * */ 

public class Payless_GUI_delayed_checkout {
	public void clickRateshopSearchBtn(ChromeDriver driver)
	{
		JavascriptExecutor jse = (JavascriptExecutor) driver;
		String clickSearchJS = "document.getElementById('searchCommandLinkResRateCode').click()";
		jse.executeScript(clickSearchJS);
	}

	private static final String NULL = null;
	private static final String RA_Num = null;
	ExtentReports extent;
	ExtentTest test;
	
	@BeforeTest
	public void startReport()
	{

		extent = Extentmanager.GetExtent();
	
	}
	@Test
	public void test() throws Exception 
	{
		try{
			Properties prop = new Properties();
			FileInputStream fis = new FileInputStream("C:\\Users\\cmn\\git\\ABG_GUI\\ABG_GUI_Automation\\src\\AVIS\\TestData\\TestDataABGGUI.properties");
			prop.load(fis);
			//WebDriver driver;
			ChromeDriver driver = new ChromeDriver();
			GUIFunctions functions = new GUIFunctions(driver);
			System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
			driver.navigate().to(prop.getProperty("PaylessURL"));
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			Thread.sleep(2000);
			functions.txt_userid.sendKeys(prop.getProperty("USERID"));
			Thread.sleep(500);
			functions.txt_password.sendKeys(prop.getProperty("PASSWORD"));
			Thread.sleep(500);
			functions.btn_login.click();
			Thread.sleep(3000);
			// Read input from excel
			///int intRowCount     = 100;
			for (int k = 1; k <= 100; k++)
			{
				Payless_GUI_delayed_checkout avis = new Payless_GUI_delayed_checkout();
				ReadWriteExcel rwe = new ReadWriteExcel("C:\\Avis_GUI_Automation\\Avis\\AVIS_GUITestData_Delay_Rental_Checkout.xlsx");
													
				String Execute = rwe.getCellData("Delayed_checkout", k, 2);
				
				//********Delete the files in the folder********//
				File file = new File(prop.getProperty("ScreenshotPayless"));  

				String[] myFiles;    
				if (file.isDirectory()) {
				    myFiles = file.list();
				    for (int i = 0; i < myFiles.length; i++) {
				        File myFile = new File(file, myFiles[i]); 
				        myFile.delete();
				    }
				}
				 int a = 28;
					
					System.out.println(" iteration " + k);
					String TCName = rwe.getCellData("Delayed_checkout", k, 4);
					String clientURL = rwe.getCellData("Delayed_checkout", k, 6);
					String outSTA = rwe.getCellData("Delayed_checkout", k, 7);
					String thinClient = clientURL + outSTA;
					//String thinClient = clientURL;
					String Resno1 = rwe.getCellData("Delayed_checkout", k, 8);
					String pswd = rwe.getCellData("Delayed_checkout", k, 9);
					String lstname = rwe.getCellData("Delayed_checkout", k, 10);
					String fstname = rwe.getCellData("Delayed_checkout", k, 11);
					String codte = rwe.getCellData("Delayed_checkout", k, 12);
					String cotme = rwe.getCellData("Delayed_checkout", k, 13);
					String insta = rwe.getCellData("Delayed_checkout", k, 14);
					String cidte = rwe.getCellData("Delayed_checkout", k, 15);                  
					String citme = rwe.getCellData("Delayed_checkout", k, 16);
					String carGrp = rwe.getCellData("Delayed_checkout", k, 17);
					String COCountry = rwe.getCellData("Delayed_checkout", k, 18);
					String COState	= rwe.getCellData("Delayed_checkout", k, 19); 
					 String CODRLICNO = rwe.getCellData("Delayed_checkout", k, 20);
					 String CODOB	= rwe.getCellData("Delayed_checkout", k, 21);
					 String COCOMPANY = rwe.getCellData("Delayed_checkout", k, 22);
					 String COADDR1	= rwe.getCellData("Delayed_checkout", k, 23);
					 String COADDR2	= rwe.getCellData("Delayed_checkout", k, 24);
					 String COADDR3	= rwe.getCellData("Delayed_checkout", k, 25);
					 String COCONTACTINFO = rwe.getCellData("Delayed_checkout", k, 26);
					 String COEMAIL	= rwe.getCellData("Delayed_checkout", k, 27);
					String awd = rwe.getCellData("Delayed_checkout", k, 28);
					String FTN = rwe.getCellData("Delayed_checkout", k, 29);
					String cardname = rwe.getCellData("Delayed_checkout", k, 30);
					String cardNo = rwe.getCellData("Delayed_checkout", k, 31);
					// System.out.print("excel card number in script :"
					// +cardNo);
					String expireMonth = rwe.getCellData("Delayed_checkout", k, 32);
					String expireYear = rwe.getCellData("Delayed_checkout", k, 33);
					String reason = rwe.getCellData("Delayed_checkout", k, 34);
					String insurance = rwe.getCellData("Delayed_checkout", k, 35);
					String counterproducts = rwe.getCellData("Delayed_checkout", k, 36);	
					String CustType = rwe.getCellData("Delayed_checkout", k, 37);
					String WizardNo = rwe.getCellData("Delayed_checkout", k, 38);	
					 String COMOP			= rwe.getCellData("Delayed_checkout", k, 39);
					 String COCCDC			= rwe.getCellData("Delayed_checkout", k, 40);
					 String CARDNO			= rwe.getCellData("Delayed_checkout", k, 41);
					 String COMM			= rwe.getCellData("Delayed_checkout", k, 42);
					 String COYY			= rwe.getCellData("Delayed_checkout", k, 43);
					 String MOPREASON		= rwe.getCellData("Delayed_checkout", k, 44);
					 String AWD				= rwe.getCellData("Delayed_checkout", k, 45);
					 String COUPON			= rwe.getCellData("Delayed_checkout", k, 46);
					 String FTNTYPE			= rwe.getCellData("Delayed_checkout", k, 47);
					 String FTNUMBER		= rwe.getCellData("Delayed_checkout", k, 48);
					 String REMARKS			= rwe.getCellData("Delayed_checkout", k, 49);
					 String COMVA			= rwe.getCellData("Delayed_checkout", k, 50);
					 String COMILEAGE		= rwe.getCellData("Delayed_checkout", k, 51);
					 String INSURANCE		= rwe.getCellData("Delayed_checkout", k, 52);
					 String COUNTERPRODUCT	= rwe.getCellData("Delayed_checkout", k, 53);
					 String ADRLASTNAME		= rwe.getCellData("Delayed_checkout", k, 54);
					 String ADRFIRSTNAME	= rwe.getCellData("Delayed_checkout", k, 55);
					 String ADRDOB			= rwe.getCellData("Delayed_checkout", k, 56);
					 String ADRCOUNTRY		= rwe.getCellData("Delayed_checkout", k, 57);
					 String ADRSTATE		= rwe.getCellData("Delayed_checkout", k, 58);
					 String ADRDRLICNO		= rwe.getCellData("Delayed_checkout", k, 59);
					 String ADREXPMM		= rwe.getCellData("Delayed_checkout", k, 60);
					 String ADREXPYY		= rwe.getCellData("Delayed_checkout", k, 61);
					 String ADRADDR1		= rwe.getCellData("Delayed_checkout", k, 62);
					 String ADRADDR2		= rwe.getCellData("Delayed_checkout", k, 63);
					 String ADRADDR3		= rwe.getCellData("Delayed_checkout", k, 64);
					 String ADRTELEPHONE1	= rwe.getCellData("Delayed_checkout", k, 65);
					 String ADRTELEPHONE2   = rwe.getCellData("Delayed_checkout", k, 66);
					 String ADRCCDC			= rwe.getCellData("Delayed_checkout", k, 67);
					 String ADRCARDNO		= rwe.getCellData("Delayed_checkout", k, 68);
					 String ADRCCEXPMM		= rwe.getCellData("Delayed_checkout", k, 69);
					 String ADRCCEXPYY		= rwe.getCellData("Delayed_checkout", k, 70);
					 String resno1		    = rwe.getCellData("Delayed_checkout", k, 71);
					 String Rentalno		= rwe.getCellData("Delayed_checkout", k, 72);
					 
					int RA_Num = Integer.parseInt(Rentalno);
					
					if (Execute.equals("Y"))
					{
						
						
						String ScreenshotPath = prop.getProperty("ScreenshotPayless");	
						
						/* Open GUI URL's */
						// System.out.println(" token URL value : " + tokenURL);
						//*******Screenshot path and test name*********//
						
						String testcasename = TCName;
						String xfilepath = prop.getProperty("ExcelPathRentalPayless") +testcasename+ ".xlsx";
						test = extent.createTest(TCName);
						
						functions.openURL(thinClient);
						/* Login */
						//functions.login(uName, pswd);
						functions.navigateToTab("ReservationRates");
						Thread.sleep(2000);				
						
						/* Enter Customer Informations */
						/* Enter FTN */
						if (rwe.getCellData("Delayed_checkout", k, 19).isEmpty()) {
							System.out.println("No FTN added");
						} else {
							functions.expandToggleBtn();
							Thread.sleep(2000);
							functions.enterFTN(FTN);
						}
						//driver.navigate().refresh();
						Thread.sleep(2000);
			
						Thread.sleep(3000);
						functions.enterCustomerName(lstname, fstname);
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						//functions.enterDriverDetail(drCountry,drState,drNumber, drDOB, drCompany,addr1,addr2, addr3, contact);
						Thread.sleep(2000);
						driver.findElement(By.id("menulist:rateshopContainer:resForm:pickupStation:pickupStation")).clear();
						Thread.sleep(2000);
						driver.findElement(By.id("menulist:rateshopContainer:resForm:pickupStation:pickupStation")).sendKeys(outSTA);
						
						 
						 //Enter the checkout date and time and checkin date and time checkout n checkin station
						functions.enterCustomerInformation(codte, cotme, insta, cidte, citme);
						functions.ScreenCapturedate(ScreenshotPath,TCName);
						//enterCustomerInformation(lstname,fstname,codte,cotme,insta,cidte,citme);
						/* Enter AWD */
						if (awd.isEmpty()) {
							System.out.println("No Avis Discount Number Added");
						} else {
							functions.enterAWD(awd);
						}

						/* Select car group */
						functions.selectCarGroupByVT(carGrp);
						Thread.sleep(2000);

						/* RATE SHOP */
						avis.clickRateshopSearchBtn(driver);
						ArrayList<WebElement> radio = (ArrayList<WebElement>) driver
								.findElements(By.xpath("//input[@name='radioRate'and @type='radio']"));
						for (int i = 0; i < radio.size(); i++) {
							if ((radio.get(i).isDisplayed()) && (radio.get(i).isEnabled())) {
								radio.get(i).click();
								if (radio.get(i).isSelected()) {
									functions.clickSelectRateBtn();
									/* Enter MOP details */
									functions.expandPaymentInfoSection();
									functions.enterPaymentInformations(cardname, cardNo, expireMonth, expireYear, reason);

									/* Add Insurances */
									Thread.sleep(5000);
									functions.expandProtectionCoverageSection();
									if (rwe.getCellData("Delayed_checkout", k, 25).isEmpty()) {
										System.out.print("No Insurance selected");
									} else {
										String[] insVal = rwe.getCellData("Delayed_checkout", k, 25).split("-");
										for (String e : insVal) {
											WebDriverWait wait1 = new WebDriverWait(driver, 10);
											if (e.equalsIgnoreCase("LDW")) {
												WebElement insurace1 = driver.findElement(
														By.id("menulist:rateshopContainer:resForm:coverageLdwYesNo"));
												Select insLDW = new Select(insurace1);
												wait1.until(ExpectedConditions.visibilityOf(insurace1));
												if (insurace1.isDisplayed()) {
													insLDW.selectByVisibleText("Yes");
												} else {
													break;
												}
											} else if (e.equalsIgnoreCase("PAI")) {
												WebElement insurace2 = driver.findElement(
														By.id("menulist:rateshopContainer:resForm:coveragePaiYesNo"));
												Select insPAI = new Select(insurace2);
												wait1.until(ExpectedConditions.visibilityOf(insurace2));
												if (insurace2.isDisplayed()) {
													insPAI.selectByVisibleText("Yes");
												} else {
													break;
												}
											} else if (e.equalsIgnoreCase("PEP")) {
												WebElement insurace3 = driver.findElement(
														By.id("menulist:rateshopContainer:resForm:coveragePepYesNo"));
												Select insPEP = new Select(insurace3);
												wait1.until(ExpectedConditions.visibilityOf(insurace3));
												if (insurace3.isDisplayed()) {
													insPEP.selectByVisibleText("Yes");
												} else {
													break;
												}
											} else if (e.equalsIgnoreCase("ALI")) {
												WebElement insurace4 = driver.findElement(
														By.id("menulist:rateshopContainer:resForm:coverageAliYesNo"));
												Select insALI = new Select(insurace4);
												wait1.until(ExpectedConditions.visibilityOf(insurace4));
												if (insurace4.isDisplayed()) {
													insALI.selectByVisibleText("Yes");
												} else {
													break;
												}
											} else if (e.equalsIgnoreCase("FSO")) {
												WebElement insurace5 = driver.findElement(
														By.id("menulist:rateshopContainer:resForm:fuelServiceOption"));
												Select insFSO = new Select(insurace5);
												wait1.until(ExpectedConditions.visibilityOf(insurace5));
												if (insurace5.isDisplayed()) {
													insFSO.selectByVisibleText("Yes");
												} else {
													break;
												}
											} else {
												break;
											}
										}
									}

									/*
									 * Add CounterProducts
									 */
									if (rwe.getCellData("Delayed_checkout", k, 26).isEmpty()) {
										System.out.print("No CounterProduct selected");
									} else {
										String[] cpVal = rwe.getCellData("Delayed_checkout", k, 26).split("-");
										for (String e : cpVal) {
											WebDriverWait wait = new WebDriverWait(driver, 10);
											try {
												if (e.equalsIgnoreCase("ADR")) {
													WebElement cp1 = driver.findElement(By.id("productQuantity40"));
													Select cpADR = new Select(cp1);
													wait.until(ExpectedConditions.visibilityOf(cp1));
													if (cp1.isDisplayed()) {
														cpADR.selectByVisibleText("1");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("CBS")) {
													WebElement cp2 = driver.findElement(By.id("productQuantity32"));
													Select cpCBS = new Select(cp2);
													wait.until(ExpectedConditions.visibilityOf(cp2));
													if (cp2.isDisplayed()) {
														cpCBS.selectByVisibleText("1");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("CSS")) {
													WebElement cp3 = driver.findElement(By.id("productQuantity34"));
													Select cpCSS = new Select(cp3);
													wait.until(ExpectedConditions.visibilityOf(cp3));
													if (cp3.isDisplayed()) {
														cpCSS.selectByVisibleText("1");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("GPS")) {
													WebElement cp4 = driver.findElement(By.id("productQuantityYesNo6"));
													Select cpGPS = new Select(cp4);
													wait.until(ExpectedConditions.visibilityOf(cp4));
													if (cp4.isDisplayed()) {
														cpGPS.selectByVisibleText("Y");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("RSN")) {
													WebElement cp5 = driver.findElement(By.id("productQuantityYesNo11"));
													Select cpRSN = new Select(cp5);
													wait.until(ExpectedConditions.visibilityOf(cp5));
													if (cp5.isDisplayed()) {
														cpRSN.selectByVisibleText("Y");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("TAB")) {
													WebElement cp6 = driver.findElement(By.id("productQuantityYesNo12"));
													Select cpTAB = new Select(cp6);
													wait.until(ExpectedConditions.visibilityOf(cp6));
													if (cp6.isDisplayed()) {
														cpTAB.selectByVisibleText("Y");
													} else {
														break;
													}
												} else if (e.equalsIgnoreCase("ESP")) {
													WebElement cp7 = driver.findElement(By.id("productQuantityYesNo6"));
													Select cpESP = new Select(cp7);
													wait.until(ExpectedConditions.visibilityOf(cp7));
													if (cp7.isDisplayed()) {
														cpESP.selectByVisibleText("Y");
													} else {
														break;
													}

												} else if (e.equalsIgnoreCase("SNB")) {
													WebElement cp8 = driver.findElement(By.id("productQuantityYesNo11"));
													Select cpSNB = new Select(cp8);
													if (cp8.isDisplayed()) {
														wait.until(ExpectedConditions.visibilityOf(cp8));
														cpSNB.selectByVisibleText("Y");
													} else {
														break;
													}
												}
											} catch (Exception e1) {
												e1.printStackTrace();
											}
										}

									}

									/*
									 * Create reservation
									 */
									functions.clickCreateReservationBtn();
									functions.ScreenCapturedate(ScreenshotPath,TCName);
									Thread.sleep(1000);
									String Resmsg = driver
											.findElement(By.xpath("//*[@id='templateInfoForm:templateInfoMsg']")).getText();
									functions.ScreenCapturedate(ScreenshotPath,TCName);
									rwe.setCellData("Delayed_checkout", k, 73, Resmsg); // write
																				// respopup
																				// in
									Thread.sleep(2000);											// excel
									String resno = Resmsg.substring(54,65);
									
									//rwe.setCellData("Delayed_co", k, 0, resno);
									rwe.setCellData("Delayed_checkout", k, 8, resno);
									Thread.sleep(1000);
									driver.findElement(By.xpath("//*[@id='templateInfoForm:templateInfoButton']")).click(); // clicks OK button in Res popup
									Thread.sleep(5000);					  
									functions.ScreenCapturedate(ScreenshotPath, TCName);
									
									//to refresh the screen
									
									
									//Click on checkout tab
									// driver.findElement(By.xpath("//input[@id='chckOutbtn']")).click();
									 driver.findElement(By.xpath("//a[@data-target='#menulist\\:checkoutlink']")).click();
									 System.out.println("Clicked on Checkout Button");
									 Thread.sleep(3000);
									 System.out.println(driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).isDisplayed());
									 driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).click();
									 Thread.sleep(2000);
									 driver.findElement(By.xpath("//input[@ng-model='checkOutSearchString']")).sendKeys(Resno1);
									 System.out.println("Entered the existing Reservation");
									 driver.findElement(By.xpath("//input[@id='searchCommandLink']")).click();
									 System.out.println("Clicked on the Search Button");
									 functions.ScreenCapturedate(ScreenshotPath,TCName);
									 Thread.sleep(20000);
									 System.out.println("Clicked on Checkout Button");
									 Thread.sleep(5000);
									 
									 //Click on cross button in checkout screen
									 System.out.println( driver.findElement(By.xpath("//div[@id='generalInfo']//div[@class='modal-content']//div[1]//span[1]")).isDisplayed());
									 Thread.sleep(2000);
									 driver.findElement(By.xpath("//div[@id='generalInfo']//div[@class='modal-content']//div[1]//span[1]")).click();
									 Thread.sleep(2000);
									 
									 //Click on Extras 
									 driver.findElement(By.xpath("//div[@class='row abg-submenu a-checkout-submenu']//span[contains(text(),'Extras')]")).click();
									 Thread.sleep(2000);
									 
									 //Click on the delayed toggled button
									 driver.findElement(By.xpath("//div[@class='a-submenu']//span[@id='checkoutToggle']")).click();
									 Thread.sleep(2000);
									 
									 //Click the reservation field and enter Resno
									 
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
									 if(FTNTYPE != NULL)
									 {
									 String ftntp= driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).getAttribute("value");
									 if (ftntp.isEmpty())
									 { 
									 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftnType']")).sendKeys(FTNTYPE);
									 System.out.println("Entered the FTN Type");
									 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:ftNumber']")).sendKeys(FTNUMBER);
									 System.out.println("Entered the FTN Number");
									 }
									 driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:verifiedFTN']")).click();
									 System.out.println("Verified the FTN number");
									 }
									 
									// driver.findElement(By.xpath("//input[@id='glyphicon glyphicon-plus']")).click();
									 
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
									 
										Select s3= new Select(driver.findElement(By.id("menulist:checkoutContainer:checkoutForm:creditCard:swipe:paymentReason")));
										s3.selectByValue(MOPREASON);
										System.out.println("Entered the Reason for payment");
										//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:mopCCI']")).sendKeys("1234");
										//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:creditCard:authorization']")).sendKeys("OK/1234");
										//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).clear();
										//driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).sendKeys("GB/C");
										String RateCode = driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:rateCode']")).getAttribute("value");
										rwe.setCellData("Delayed_checkout", k, 74,RateCode);
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mvaOrParkingSpace']")).sendKeys(COMVA);
										System.out.println("Entered the MVA number");
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).click();
										Thread.sleep(2000);
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).clear();
										Thread.sleep(2000);
										driver.findElement(By.xpath("//input[@id='menulist:checkoutContainer:checkoutForm:mileage']")).sendKeys(COMILEAGE);
										System.out.println("Entered the Mileage");
										functions.ScreenCapturedate(ScreenshotPath,TCName);
										//Click on Delay continue button
										if(driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).isDisplayed())
										{
										driver.findElement(By.xpath("//input[@id='footerForm:continueVehicleDelayButton']")).click();
										}
										Thread.sleep(10000);
										if(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).isDisplayed())
										{
										String OptMSG= driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getAttribute("value");
										System.out.println("The Output message in vechicle continue screen is "+ OptMSG);
										if(OptMSG.equals("NET RATE ERROR")) 
										{
										System.out.println("ERROR MESSAGE "+OptMSG+" is Displayed");
										rwe.setCellData("Delayed_checkout", k, 82, OptMSG);
										rwe.setCellData("Delayed_checkout", k, 83, "FAIL");
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
											rwe.setCellData("Delayed_checkout", k, 82, OptMSG);
											rwe.setCellData("Delayed_checkout", k, 83, "FAIL");
											test.log(Status.FAIL, "Fail");
											
											functions.ScreenCapturedate(ScreenshotPath,TCName);
											System.out.println("The Screenshot is taken");
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptCloseButton")).click();
											driver.close();
											break;	
												
											}
										}
										
										//Enter Rental agreement number
										Thread.sleep(5000);
										//CharSequence[] RA_Num = get_RA_Number(Integer.parseInt(Rentalno));
										//driver.findElement(By.xpath("//*[@class='form-control allowBackspace ng-touched']")).click();
										//driver.findElement(By.xpath("//*[@class='form-control allowBackspace ng-touched']")).sendKeys(Rentalno);
										//driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
										//driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(Rentalno);
										
										//Thread.sleep(5000);
										functions.ScreenCapturedate(ScreenshotPath,TCName);
										/*if (driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).isDisplayed())
										{											
					 				  	driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
					 				  	System.out.println("Clicked on Continue Button");
					 				  	
					 				  	
										if (documentnoinuse.contains("DOCUMENT NUMBER ALREADY IN USE")) {*/
											/*while(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText().contains("C/O CAR INFO DISCREPANCIES") == false) {

												
												//String new_Ra2 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
												//String RA_num1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
												//driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(RA_num1);
												int new_RA = get_RA_Number(RA_Num);
												
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(Integer.toString(new_RA));
												Thread.sleep(2000);
												String RA_num1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
												Thread.sleep(2000);
												//System.out.println(RA_num1);
												Thread.sleep(2000);
												rwe.setCellData("Delayed_checkout", k, 72, RA_num1);
												Thread.sleep(2000);
												driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
												Thread.sleep(5000);

												}*/
											
											do {
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(Rentalno);
												driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
							 				  	System.out.println("Clicked on Continue Button");
							 				  	String documentnoinuse = driver.findElement(By.xpath("//*[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
												if (documentnoinuse.contains("DOCUMENT NUMBER ALREADY IN USE")) {
												int new_RA = get_RA_Number(RA_Num);
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
												driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(Integer.toString(new_RA));
												Thread.sleep(2000);
												String RA_num1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).getAttribute("value");
												Thread.sleep(2000);
												System.out.println(RA_num1);
												Thread.sleep(2000);
												rwe.setCellData("Delayed_checkout", k, 72, RA_num1);
												Thread.sleep(2000);
												driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
												Thread.sleep(5000);
												}
											}while(driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText().contains("C/O CAR INFO DISCREPANCIES") == false);
											
					 				  	Thread.sleep(10000);
										functions.ScreenCapturedate(ScreenshotPath,TCName);
										// For message as C/O CAR INFO DISCREPANCIES
										String afterRa1 = driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptOut")).getText();
										if (afterRa1.contains("C/O CAR INFO DISCREPANCIES")) or (afterRa1.contains("DO NOT RENT THIS CAR"));
										{
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).click();
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).clear();
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys(COMILEAGE);
											Thread.sleep(2000);
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:1:repromptValue")).click();
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:1:repromptValue")).clear();
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:1:repromptValue")).sendKeys(COMVA);
											Thread.sleep(2000);
											driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
											Thread.sleep(10000);
										}
										if(driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).isDisplayed())
										{
										driver.findElement(By.xpath("//input[@id='footerForm:completeCheckoutButton']")).click();
										System.out.println("Clicked on Complete Checkout Button");
										Thread.sleep(7000);
										}
										
										Thread.sleep(7000);
										try
										{
											//functions.ScreenCapturedate(ScreenshotPath,TCName);
											 if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())
													 
											 {
												String str1= driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
												System.out.println("The message displayed is "+str1);
												if (str1.contains("Please enter the Credit ID security code (CVV/CCV)"))
												{
													if(COCCDC.equals("CA"))
										               {
										                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("1234");
										                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
										                    Thread.sleep(12000);
										                    
										               }
										               else
										               {
										                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue")).sendKeys("123");
										                    driver.findElement(By.id("checkoutRepromptDialog:repromptForm:repromptSubmitButton")).click();
										                    Thread.sleep(12000);
										                    
										               }
													Thread.sleep(5000);
													 if (driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())								                                             
													 
													 {
															String str2=driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
															System.out.println("The message displayed is "+str2); 
															if(str2.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION"))
																              
																	{
																 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
																 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
																 driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
																 Thread.sleep(12000);
																	}
														 }
													
													
												}
												else
													functions.ScreenCapturedate(ScreenshotPath,TCName);
													if (str1.contains("**MULTIPLE RENTALS**NEEDS MANAGEMENT AUTHORIZATION"))
												{
													driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).click();
													driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValuePassword']")).sendKeys("YES");
													driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
													Thread.sleep(6000);
												}
										}
											 
										
										if(driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).isDisplayed())
										{
										String str2=driver.findElement(By.xpath("//textarea[@id='checkoutRepromptDialog:repromptForm:repromptOut']")).getText();
										System.out.println("The message displayed is "+str2); 
										if(str2.contains("DO NOT RENT/CALL APPROPRIATE"))
		              
										{
											 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue']")).click();
											 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:0:repromptValue']")).sendKeys("Ok/1234");
											 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValueSWR']")).click();
											 driver.findElement(By.xpath("//input[@id='checkoutRepromptDialog:repromptForm:repromptTable:2:repromptValueSWR']")).sendKeys("R");
											 driver.findElement(By.xpath("//button[@id='checkoutRepromptDialog:repromptForm:repromptSubmitButton']")).click();
											 Thread.sleep(12000);
										}
										
										}
										}
										catch (Exception e)
										{
											e.printStackTrace();
										}
										}
										functions.ScreenCapturedate(ScreenshotPath,TCName);
										//download the rental pdf
										Thread.sleep(2000);
										//driver.findElement(By.xpath("//span[@class='ui-button-text ui-c'][contains(text(),'Close')]")).click();
										
										      
									 // Checkout complete Screen and update in Excel 
										   
										  //tring screenname = driver.findElement(By.xpath("//h5[@class='completeCheckout ng-binding']")).getText();
												   
											Thread.sleep(7000);			                    
										   //String strCOLNFNGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteName")).getText();
										   String strCOLNFNGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteName']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 75, strCOLNFNGetText);
							               System.out.println(" Last Name First Name value is " + strCOLNFNGetText);
							              // String strCOVehicleModelGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle")).getText();
							               String strCOVehicleModelGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteVehicle']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 76, strCOVehicleModelGetText);
							               System.out.println(" Vehicle Model value is " + strCOVehicleModelGetText);
							               //String strCOResNoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum")).getText();
							               String strCOResNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteResNum']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 77, strCOResNoGetText);
							               System.out.println(" Reservation No value is " + strCOResNoGetText);
							               //String strCOMVAGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA")).getText();
							               String strCOMVAGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteMVA']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 78, strCOMVAGetText);
							               System.out.println(" MVA No value is " + strCOMVAGetText);
							               //String strCORANoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber")).getText();
							               String strCORANoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteRANumber']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 79, strCORANoGetText);
							               System.out.println(" RA Number value is " + strCORANoGetText);
							              // String strCOLicensePlateNoGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate")).getText();
							               String strCOLicensePlateNoGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteLicensePlate']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 80, strCOLicensePlateNoGetText);
							               System.out.println(" License Plate Number value is " + strCOLicensePlateNoGetText);
							               //String strCOEstimatedCostGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost")).getText();
							               String strCOEstimatedCostGetText = driver.findElement(By.xpath("//div[@class='form-group']//span[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteEstimatedCost']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 81, strCOEstimatedCostGetText);
							               System.out.println(" Estimated Cost value is " + strCOEstimatedCostGetText);
							               //String strCOSystemMsgGetText = driver.findElement(By.id("checkoutDialog:checkoutCompleteForm:checkoutCompleteOut")).getText();
							               String strCOSystemMsgGetText = driver.findElement(By.xpath("//div[@class='form-group']//textarea[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteOut']")).getText();
							               Thread.sleep(7000);
							               rwe.setCellData("Delayed_checkout", k, 82, strCOSystemMsgGetText);
							               System.out.println(" CheckOut Complete System Message value is " + strCOSystemMsgGetText);
							               Thread.sleep(7000);
							               
										
										
							              // functions.ScreenCapturedate(ScreenshotPath,TCName);
										System.out.println("The Screenshot is taken");	
										Thread.sleep(5000);
										if (driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).isDisplayed())
							               {
											driver.findElement(By.xpath("//input[@id='checkoutDialog:checkoutCompleteForm:checkoutCompleteDoneButton']")).click();
							               System.out.println("Clicked on Done Button");
							               Thread.sleep(7000);
							               }
										//String ScreenShotPath2 = "C:\\Selenium\\Screenshots\\Avis\\Avis_CreateRentals\\";
										//functions.ScreenCapturedate(ScreenshotPath,TCName);
										System.out.println("The Screenshot is taken");
										//driver.findElement(By.xpath("//input[@id='headerLogOutButton']")).click();
										//driver.findElement(By.xpath("//button[@id='logoutForm:yesLogoutButton']")).click();
										//driver.close();
										Thread.sleep(2000);
										//functions.closeWindows();
										//test = extent.createTest(TCName);	
					                    if(rwe.getCellData("Delayed_checkout", k, 82).isEmpty())
					                    {
											rwe.setCellData("Delayed_checkout", k, 83,"FAIL");
									test.log(Status.FAIL, "Fail");
										} else {
											rwe.setCellData("Delayed_checkout", k, 83, "PASS");
									test.log(Status.PASS, "Pass");
										}
									 
									 functions.ScreenCapturedate(ScreenshotPath,TCName);
									 Thread.sleep(20000);
									
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
									       File f = new File(prop.getProperty("ScreenshotPayless"));

									       // Populates the array with names of files and directories
									       pathnames = f.list();

									       // For each pathname in the pathnames array
									       for (String pathname : pathnames) {
									             // Print the names of files and directories

									             InputStream inputStream = new FileInputStream(
									            		 prop.getProperty("ScreenshotPayless")+pathname);
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
								} 
							

							else {
								System.out.println("Execution status is N for iteration " + k + "...");
							}
						}
						
} //End of if statement
					
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
					
					
				

private void or(boolean contains) {
	// TODO Auto-generated method stub
	
}

public static int get_RA_Number(int RA_Num1) {

	if (Integer.toString((RA_Num1)).endsWith("6")) {
		RA_Num1 = RA_Num1 + 4;
		System.out.println(RA_Num1);
	}

	else {
		RA_Num1 = RA_Num1 + 11;
		//System.out.println(RA_Num);
	}
	return RA_Num1;
	
	
}
}
