package com.sn.ajr.citation;
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets;
import com.sn.ajr.ProjGSheet;

import org.testng.annotations.BeforeTest;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
//import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
public class Awos {
	public static WebDriver driver;
	@Test( timeOut = 300000 )
  public static void classname() throws IOException, GeneralSecurityException, InterruptedException {
	  System.out.println("wos");
	  File currentDirFile = new File(".");
		String rootpath1 = currentDirFile.getAbsolutePath();
		String rootpath = rootpath1.replace(".", "");
		String downloadFilepath = rootpath + "DownloadingTempfiles\\";
		String livepath = rootpath + "batchfiles\\live.txt"; 
		String sheetidq = "";
  	File fa1 = new File(livepath);
  	if(fa1.exists()){
  		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
  	}else {
  		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
  	}
//	  String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example");   
	  	String JournalName = "";
	  	String wosid = "";
	  	String wosp = "";
	  	try {
	  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
			for (List<?> row : values) {
				JournalName = (String) row.get(2); 
			}		
			List<List<Object>> values1 = sheetq.getCellData(sheetsService1, "Details!C5:E5",sheetidq); 
			for (List<?> row : values1) {
				wosid = (String) row.get(0);
				wosp = (String) row.get(2);
			}
	  	} catch (Exception e1) {}		
	  	if(!JournalName.contentEquals("")) {
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--disable-gpu");
		String path1 = "C:/Users/bhushanhadmin/AppData/Local/Google/Chrome/User Data/Work";
		File fa0 = new File(path1);
    	if(fa0.exists()){
    		options.addArguments("user-data-dir=C:/Users/bhushanhadmin/AppData/Local/Google/Chrome/User Data/Work");
		}else {
    		options.addArguments("user-data-dir=C:/Users/hake01/AppData/Local/Google/Chrome/User Data/Profile 2");	
    	}
		driver = new ChromeDriver(options);
		
//web of science
try {
for (int a = 1; a < 5; a++) {
driver.get("https://www.webofscience.com/wos/woscc/basic-search");
//driver.get("https://access.clarivate.com/login?app=wos");
driver.manage().window().maximize();
Thread.sleep(1000); 
try{
new WebDriverWait(driver, Duration.ofSeconds(30))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='email']")));
if(!driver.findElements(By.xpath("//*[@name='email']")).isEmpty()){
	driver.findElement(By.xpath("//*[@name='email']")).sendKeys(wosid);
	driver.findElement(By.xpath("//*[@name='password']")).sendKeys(wosp);
driver.findElement(By.id("signIn-btn")).click();
Thread.sleep(5000);
}
} catch(Exception e1) {}
try {
if(!driver.findElements(By.xpath("//*[@id='onetrust-accept-btn-handler']")).isEmpty()){
driver.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']")).click();
}
if(!driver.findElements(By.xpath("//*[@class='NEWwokErrorContainer SignInLeftColumn']//a[1]")).isEmpty()){
driver.findElement(By.xpath("//*[@class='NEWwokErrorContainer SignInLeftColumn']//a[1]")).click();
Thread.sleep(10000);
}
if(!driver.findElements(By.xpath("//*[@aria-label='Close this dialog']//mat-icon")).isEmpty()){ 
driver.findElement(By.xpath("//*[@aria-label='Close this dialog']//mat-icon")).click();
}
} catch(Exception e1) {}
try {
WebElement searchlable = new WebDriverWait(driver, Duration.ofSeconds(30))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='snSearchType']//button")));
JavascriptExecutor jsq = (JavascriptExecutor) driver;

try {
if(!driver.findElements(By.xpath("//*[@id='onetrust-accept-btn-handler']")).isEmpty()){ 
driver.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']")).click();
}
} catch(Exception e1) {}
jsq.executeScript("arguments[0].scrollIntoView(true);", searchlable);
searchlable.click();
} catch(Exception e1) {}
WebElement searchselect = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='options']//div[5]")));
searchselect.click();
try {
if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
driver.findElement(By.className("_pendo-close-guide")).click();
}
} catch(Exception e1) {}
Thread.sleep(100);
WebElement jtitlesear = new WebDriverWait(driver, Duration.ofSeconds(20))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Search box 1']")));

try {
driver.findElement(By.xpath("//*[@aria-label='Add row']")).click();
} catch(Exception e1) {
driver.findElement(By.xpath("(//*[@data-mat-icon-name='add'])[1]")).click();
}
Thread.sleep(500);
driver.findElement(By.xpath("(//*[@aria-label='Select search field All Fields'])[1]")).click();
Thread.sleep(500);
WebElement searchselect1 = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='options']//div[6]")));
searchselect1.click();
Thread.sleep(500);
driver.findElement(By.xpath("//*[@aria-label='Search box 2']")).clear();
driver.findElement(By.xpath("//*[@aria-label='Search box 2']")).sendKeys("2016-2020");
Thread.sleep(500);
jtitlesear.clear();
Thread.sleep(500);
jtitlesear.sendKeys(JournalName);
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@data-ta=\"basic-search-link\"]")).click();
WebElement runsearch = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='run-search']//span")));
JavascriptExecutor jsa = (JavascriptExecutor) driver;
jsa.executeScript("arguments[0].scrollIntoView(true);", runsearch);
runsearch.click();
//Thread.sleep(15000);

try {
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='snSearchType']//b")));
} catch(Exception e1) {System.out.println("Exception");}
if(!driver.findElements(By.xpath("//*[@id='snSearchType']//b")).isEmpty()){
String checkstatus = driver.findElement(By.xpath("//*[@id='snSearchType']//b")).getText();

if (checkstatus.contentEquals("Your search found no results")) {
if(JournalName.startsWith("The ")){
	JournalName = JournalName.replace("The ", "");
} 
if(JournalName.startsWith("Der ")){
	JournalName = JournalName.replace("Der ", "");
}
WebElement jtitlesearch = new WebDriverWait(driver, Duration.ofSeconds(100))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='search-main-box']")));
jtitlesearch.clear();
Thread.sleep(500);
jtitlesearch.sendKeys(JournalName);
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@data-ta=\"basic-search-link\"]")).click();
WebElement runsearch1 = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='run-search']//span")));
JavascriptExecutor jsa1 = (JavascriptExecutor) driver;
jsa1.executeScript("arguments[0].scrollIntoView(true);", runsearch1);
runsearch.click();

try {
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='snSearchType']//b")));
} catch(Exception e1) {System.out.println("Exception");}
if(!driver.findElements(By.xpath("//*[@id='snSearchType']//b")).isEmpty()){ 
String checkstatus1 = driver.findElement(By.xpath("//*[@id='snSearchType']//b")).getText();
if (checkstatus1.contentEquals("Your search found no results")) {
	if(JournalName.indexOf("&") > 0){
		JournalName = JournalName.replace("&", "and");
	}
WebElement jtitlesearch1 = new WebDriverWait(driver, Duration.ofSeconds(100))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='search-main-box']")));
jtitlesearch1.clear();
Thread.sleep(500);
jtitlesearch1.sendKeys(JournalName);
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@data-ta=\"basic-search-link\"]")).click();
WebElement runsearch2 = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='run-search']//span")));
JavascriptExecutor jsa2 = (JavascriptExecutor) driver;
jsa2.executeScript("arguments[0].scrollIntoView(true);", runsearch2);
runsearch2.click(); 

try {
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='snSearchType']//b")));
} catch(Exception e1) {System.out.println("Exception");}
if(!driver.findElements(By.xpath("//*[@id='snSearchType']//b")).isEmpty()){ 
String checkstatus2 = driver.findElement(By.xpath("//*[@id='snSearchType']//b")).getText();
if (checkstatus2.contentEquals("Your search found no results")) {
	if(JournalName.indexOf("and") > 0){
		JournalName = JournalName.replace("and", "&");
	}
WebElement jtitlesearch2 = new WebDriverWait(driver, Duration.ofSeconds(100))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='search-main-box']")));
jtitlesearch2.clear();
Thread.sleep(500);
jtitlesearch2.sendKeys(JournalName);
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@data-ta=\"basic-search-link\"]")).click();
WebElement runsearch3 = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='run-search']//span")));
JavascriptExecutor jsa3 = (JavascriptExecutor) driver;
jsa3.executeScript("arguments[0].scrollIntoView(true);", runsearch3);
runsearch2.click();

try {
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='snSearchType']//b")));
} catch(Exception e1) {System.out.println("Exception");}
if(!driver.findElements(By.xpath("//*[@id='snSearchType']//b")).isEmpty()){ 
String checkstatus4 = driver.findElement(By.xpath("//*[@id='snSearchType']//b")).getText();
if (checkstatus4.contentEquals("Your search found no results")) {
driver.close();
driver.quit();
}
}
} 
}
}
}}}

Thread.sleep(2000);
try {
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='marked-list-link']")));
driver.findElement(By.xpath("//*[@data-ta='marked-list-link']")).click(); 
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='unfiled-mls-data-container']//tbody//tr[1]//td//a")));
//Thread.sleep(5000);
driver.findElement(By.xpath("//*[@class='unfiled-mls-data-container']//tbody//tr[1]//td//a")).click();

try {
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='remove-from-marked-list']")));
driver.findElement(By.xpath("//*[@data-ta='remove-from-marked-list']")).click(); 
} catch(Exception e1) {System.out.println("Exception");}
Thread.sleep(100);
driver.findElement(By.xpath("(//*[@value=\"fromWholeList\"])[1]")).click(); 
Thread.sleep(100);
driver.findElement(By.xpath("(//*[@data-ta=\"submit-remove-from-marked-list\"])[1]")).click(); 
Thread.sleep(100);
driver.findElement(By.xpath("//*[@class=\"cdk-overlay-pane\"]//button[2]")).click();
} catch(Exception e1) {System.out.println("Exception");}
Thread.sleep(100);
driver.findElement(By.xpath("//*[@id=\"snHeaderLink\"]")).click(); 
driver.findElement(By.xpath("//*[@data-ta='run-search']//span")).click(); 
try {
if(!driver.findElements(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).isEmpty()){ 
driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).click();
}
} catch(Exception e1) {System.out.println("Exception");}
if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
driver.findElement(By.className("_pendo-close-guide")).click();
}

//check
Thread.sleep(3000);
if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
	driver.findElement(By.className("_pendo-close-guide")).click();
}
if(!driver.findElements(By.xpath("//*[@aria-label='Close this tour']")).isEmpty()){ 
	driver.findElement(By.xpath("//*[@aria-label='Close this tour']")).click(); 
}
//driver.findElement(By.xpath("//*[@title='Publication Years: 2019']")).click(); 
if(!driver.findElements(By.xpath("//*[@aria-label='See all Document Types']")).isEmpty()){
driver.findElement(By.xpath("//*[@aria-label='See all Document Types']")).click(); 
}
Thread.sleep(4000);
if(!driver.findElements(By.xpath("//*[@title='Document Types: Articles']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Articles']")).click(); 
}
if(!driver.findElements(By.xpath("//*[@title='Document Types: Article']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Article']")).click(); 
}
if(!driver.findElements(By.xpath("//*[@title='Document Types: Review Articles']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Review Articles']")).click(); 
}
if(!driver.findElements(By.xpath("//*[@title='Document Types: Review Article']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Review Article']")).click(); 
}
if(!driver.findElements(By.xpath("//*[@title='Document Types: Proceedings Papers']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Proceedings Papers']")).click();
}
if(!driver.findElements(By.xpath("//*[@title='Document Types: Proceeding Paper']")).isEmpty()){
driver.findElement(By.xpath("//*[@title='Document Types: Proceeding Paper']")).click();
}
if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
driver.findElement(By.className("_pendo-close-guide")).click();
}
if(!driver.findElements(By.xpath("//*[@class=\"seeAll-footer\"]//button[3]")).isEmpty()){
	driver.findElement(By.xpath("//*[@class=\"seeAll-footer\"]//button[3]")).click();
}else {
if(!driver.findElements(By.xpath("(//*[@data-ta='refine-submit'])[3]")).isEmpty()){
	driver.findElement(By.xpath("(//*[@data-ta='refine-submit'])[3]")).click();
}
}
//check
Thread.sleep(1000);
if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
driver.findElement(By.className("_pendo-close-guide")).click();
}

String crvalue = driver.findElement(By.xpath("(//*[@class=\"search-display\"]//span)[1]")).getText();
crvalue = crvalue.replace(",", "");
int crvaluem=Integer.parseInt(crvalue); 

//marklist
driver.findElement(By.xpath("//*[@data-ta=\"add-to-marked-list\"]")).click();
Thread.sleep(100);
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"From input\"]")).sendKeys("1");
driver.findElement(By.xpath("//*[@aria-label=\"To input\"]")).sendKeys("50000");
driver.findElement(By.xpath("//*[@aria-label=\"Add\"]")).click();
Thread.sleep(5000);
new WebDriverWait(driver, Duration.ofSeconds(50))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='marked-list-link']")));
driver.findElement(By.xpath("//*[@data-ta='marked-list-link']")).click();
Thread.sleep(2000);
if (crvaluem > 1 ) {
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='marked-list-link']")));
	driver.findElement(By.xpath("//*[@data-ta='marked-list-link']")).click(); 
	Thread.sleep(1000);
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='unfiled-mls-data-container']//tbody//tr[1]//td//a")));	
	driver.findElement(By.xpath("//*[@class='unfiled-mls-data-container']//tbody//tr[1]//td//a")).click();
	Thread.sleep(5000);
driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
Thread.sleep(50);
driver.findElement(By.xpath("//*[@id='exportToExcelButton']")).click();
//Thread.sleep(2000);
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@value='fromRange']//span")));
driver.findElement(By.xpath("//*[@value='fromRange']//span")).click();
driver.findElement(By.xpath("//*[@name='markFrom']")).clear();
driver.findElement(By.xpath("//*[@name='markTo']")).clear();
driver.findElement(By.xpath("//*[@name='markFrom']")).sendKeys("1");
driver.findElement(By.xpath("//*[@name='markTo']")).sendKeys("1000");
driver.findElement(By.xpath("//*[@aria-label=' Author, Title, Source']")).click();
driver.findElement(By.xpath("//*[@title='Full Record']")).click();
driver.findElement(By.xpath("(//app-export-out-details//button)[3]")).click();
}
Thread.sleep(1000);

if (crvaluem > 1000 ) {
try {	
driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
} catch(Exception e1) {
	Thread.sleep(7000);
	try {
	driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
	} catch(Exception e2) {}
}
Thread.sleep(50);
driver.findElement(By.xpath("//*[@id='exportToExcelButton']")).click(); 
new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@value='fromRange']//span")));
driver.findElement(By.xpath("//*[@value='fromRange']//span")).click();
driver.findElement(By.xpath("//*[@name='markFrom']")).clear();
driver.findElement(By.xpath("//*[@name='markTo']")).clear();
driver.findElement(By.xpath("//*[@name='markFrom']")).sendKeys("1001");
driver.findElement(By.xpath("//*[@name='markTo']")).sendKeys("2000");
driver.findElement(By.xpath("//*[@aria-label=' Author, Title, Source']")).click();
driver.findElement(By.xpath("//*[@title='Full Record']")).click();
driver.findElement(By.xpath("(//app-export-out-details//button)[3]")).click();
}
Thread.sleep(5000);

if (crvaluem > 2000 ) {
	try {
		driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
		} catch(Exception e1) {
			Thread.sleep(7000);
			try {
			driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
			} catch(Exception e2) {}
		}
		Thread.sleep(50);
		driver.findElement(By.xpath("//*[@id='exportToExcelButton']")).click();
Thread.sleep(1500);
driver.findElement(By.xpath("//*[@value='fromRange']//span")).click();
driver.findElement(By.xpath("//*[@name='markFrom']")).clear();
driver.findElement(By.xpath("//*[@name='markTo']")).clear();
driver.findElement(By.xpath("//*[@name='markFrom']")).sendKeys("2001");
driver.findElement(By.xpath("//*[@name='markTo']")).sendKeys("3000");
driver.findElement(By.xpath("//*[@aria-label=' Author, Title, Source']")).click();
driver.findElement(By.xpath("//*[@title='Full Record']")).click();
driver.findElement(By.xpath("(//app-export-out-details//button)[3]")).click();
}
Thread.sleep(5000);

if (crvaluem > 3000 ) {
	try {
		driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
		} catch(Exception e1) {
			Thread.sleep(7000);
			try {
			driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
			} catch(Exception e2) {}
		}
		Thread.sleep(50);
		driver.findElement(By.xpath("//*[@id='exportToExcelButton']")).click(); 
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@value='fromRange']//span")).click();
driver.findElement(By.xpath("//*[@name='markFrom']")).clear();
driver.findElement(By.xpath("//*[@name='markTo']")).clear();
driver.findElement(By.xpath("//*[@name='markFrom']")).sendKeys("3001");
driver.findElement(By.xpath("//*[@name='markTo']")).sendKeys("4000");
driver.findElement(By.xpath("//*[@aria-label=' Author, Title, Source']")).click();
driver.findElement(By.xpath("//*[@title='Full Record']")).click();
driver.findElement(By.xpath("(//app-export-out-details//button)[3]")).click();
}
Thread.sleep(5000);

if (crvaluem > 4000 ) {
	try {
		driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
		} catch(Exception e1) {
			Thread.sleep(7000);
			try {
			driver.findElement(By.xpath("//*[@aria-label='summaryRecordsTop']//app-export-menu")).click();
			} catch(Exception e2) {}
		}
		Thread.sleep(50);
		driver.findElement(By.xpath("//*[@id='exportToExcelButton']")).click(); 
Thread.sleep(1000);
driver.findElement(By.xpath("//*[@value='fromRange']//span")).click();
driver.findElement(By.xpath("//*[@name='markFrom']")).clear();
driver.findElement(By.xpath("//*[@name='markTo']")).clear();
driver.findElement(By.xpath("//*[@name='markFrom']")).sendKeys("4001");
driver.findElement(By.xpath("//*[@name='markTo']")).sendKeys("5000");
driver.findElement(By.xpath("//*[@aria-label=' Author, Title, Source']")).click();
driver.findElement(By.xpath("//*[@title='Full Record']")).click();
driver.findElement(By.xpath("(//app-export-out-details//button)[3]")).click();
}

//next page 
WebElement citationr = new WebDriverWait(driver, Duration.ofSeconds(200))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-ta='run-citation-report']")));
citationr.click(); 
Thread.sleep(3000);

if (crvaluem > 1 ) {
WebElement exreport = new WebDriverWait(driver, Duration.ofSeconds(20))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Export Full Report']")));
exreport.click();
WebElement filex = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='fileName']")));
filex.clear(); 
filex.sendKeys("citaion1"); 
WebElement cateee=driver.findElement(By.xpath("//*[@name='fileName']")); 
JavascriptExecutor jsae = (JavascriptExecutor) driver;
jsae.executeScript("arguments[0].scrollIntoView(true);", cateee); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[1]")).click();
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).clear();
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).clear();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).sendKeys("1");
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).sendKeys("1000"); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[3]")).click();
//WebElement cateee1 = driver.findElement(By.xpath("//*[@class='button-row']//button"));  
//jsae.executeScript("arguments[0].scrollIntoView(true);", cateee1); 
driver.findElement(By.xpath("//*[@class='button-row']//button")).click();
}
Thread.sleep(5000);
if (crvaluem > 1000 ) {
WebElement exreport = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Export Full Report']")));
exreport.click();
WebElement filex = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='fileName']")));
filex.clear();
filex.sendKeys("citaion2"); 
WebElement cateee=driver.findElement(By.xpath("//*[@name='fileName']")); 
JavascriptExecutor jsaeq = (JavascriptExecutor) driver;
jsaeq.executeScript("arguments[0].scrollIntoView(true);", cateee); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[1]")).click();
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).clear();
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).clear();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).sendKeys("1001");
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).sendKeys("2000"); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[3]")).click();
//WebElement cateee2 = driver.findElement(By.xpath("//*[@class='button-row']//button"));  
//jsaeq.executeScript("arguments[0].scrollIntoView(true);", cateee2); 
driver.findElement(By.xpath("//*[@class='button-row']//button")).click();
}
Thread.sleep(5000);
if (crvaluem > 2000 ) {
WebElement exreport = new WebDriverWait(driver, Duration.ofSeconds(20))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Export Full Report']")));
exreport.click();
WebElement filex = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='fileName']")));
filex.clear();
filex.sendKeys("citaion3"); 
WebElement cateee=driver.findElement(By.xpath("//*[@name='fileName']")); 
JavascriptExecutor jsaew = (JavascriptExecutor) driver;
jsaew.executeScript("arguments[0].scrollIntoView(true);", cateee); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[1]")).click();
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).clear();
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).clear();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).sendKeys("2001");
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).sendKeys("3000");
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[3]")).click();
driver.findElement(By.xpath("//*[@class='button-row']//button")).click();
}
Thread.sleep(5000);
if (crvaluem > 3000 ) {
WebElement exreport = new WebDriverWait(driver, Duration.ofSeconds(20))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Export Full Report']")));
exreport.click();
WebElement filex = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='fileName']")));
filex.sendKeys("citaion4"); 
WebElement cateee=driver.findElement(By.xpath("//*[@name='fileName']")); 
JavascriptExecutor jsaee = (JavascriptExecutor) driver;
jsaee.executeScript("arguments[0].scrollIntoView(true);", cateee); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[1]")).click();
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).clear();
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).clear();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).sendKeys("3001");
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).sendKeys("4000");
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[3]")).click();
driver.findElement(By.xpath("//*[@class='button-row']//button")).click();
}
Thread.sleep(5000);
if (crvaluem > 4000 ) {
WebElement exreport = new WebDriverWait(driver, Duration.ofSeconds(20))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@aria-label='Export Full Report']")));
exreport.click();
WebElement filex = new WebDriverWait(driver, Duration.ofSeconds(10))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='fileName']")));
filex.clear();
filex.sendKeys("citaion5"); 
WebElement cateee=driver.findElement(By.xpath("//*[@name='fileName']")); 
JavascriptExecutor jsae = (JavascriptExecutor) driver;
jsae.executeScript("arguments[0].scrollIntoView(true);", cateee); 
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[1]")).click();
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[2]")).click();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).clear();
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).clear();
driver.findElement(By.xpath("//*[@aria-label=\"Input starting record range\"]")).sendKeys("4001");
driver.findElement(By.xpath("//*[@aria-label='Input ending record range. A maximum of {{total}} records can be exported at one time.']")).sendKeys("5000");
driver.findElement(By.xpath("(//*[@class=\"mat-radio-label-content\"])[3]")).click();
driver.findElement(By.xpath("//*[@class='button-row']//button")).click();
}
Thread.sleep(5000);

String catsheet = downloadFilepath + "citaion1";
File f0 = new File(catsheet + ".xls");
if(f0.exists()){
	a = 5;
}
}
driver.close();
driver.quit();
System.out.println("Done");
} catch(Exception e1) {}
  	try {
  		driver.close();
  		driver.quit();
  		} catch(Exception e1) {}}
  }

  @BeforeTest
  public void beforeTest() {
  }

  @AfterTest
  public void afterTest() {	 
	  try {
			driver.close();
			driver.quit();
		} catch(Exception e1) {System.out.println("Exceptionbex");}
  }

}
