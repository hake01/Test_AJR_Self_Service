package com.sn.ajr.googledata;
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.slides.v1.Slides;
import com.sn.ajr.ProjGSheet;
import com.sn.ajr.ProjGSlide;
import com.sn.ajr.UpdateSlidePojo;
import org.testng.annotations.BeforeTest;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap; 
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
public class GoogleData {
	public static WebDriver driver;
  @Test( timeOut = 350000 )
  public static void classname() throws IOException, GeneralSecurityException, InterruptedException {
	  	System.out.println("googledata start");
	  	File currentDirFile = new File(".");
		String rootpath1 = currentDirFile.getAbsolutePath();
		String rootpath = rootpath1.replace(".", "");
		String downloadFilepath = rootpath + "DownloadingTempfiles\\";
//		downloadFilepath = downloadFilepath.replace("\\","/");
//		System.out.println("downloadFilepath" + downloadFilepath);
//		System.out.println("downloadFilepath" + path1);
		String livepath = rootpath + "batchfiles\\live.txt"; 
		String profile = "";
		
		String path1 = "C:/Users/bhushanhadmin/AppData/Local/Google/Chrome/User Data/Work";
		File fa0 = new File(path1);
		
		String sheetidq = "";
    	File f0 = new File(livepath);
    	if(f0.exists()){
    		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
    		if(fa0.exists()){
        		profile = "user-data-dir=C:/Users/bhushanhadmin/AppData/Local/Google/Chrome/User Data/Work1";
        		}
        		else {
        		profile = "user-data-dir=C:/Users/hake01/AppData/Local/Google/Chrome/User Data/Work1";
        		}
    	}else {
    		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
    		if(fa0.exists()){
    		profile = "user-data-dir=C:/Users/bhushanhadmin/AppData/Local/Google/Chrome/User Data/Work";
    		}
    		else {
    		profile = "user-data-dir=C:/Users/hake01/AppData/Local/Google/Chrome/User Data/Work";
    		}
    	}
    	
//	  	String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example"); 
		String sheetname = "";
    	String JournalID = "";
    	String JournalName = "";
    	String Platform = ""; 
    	String pptid = "";
        String gid = "";
        String gidp = "";
        String gide ="";
        String sheetid = "";
        for (int a = 1; a < 5; a++) {
        	try {
    	try {
        	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
    		for (List<?> row : values) {
    			sheetname = (String) row.get(0);
    	    	JournalID = (String) row.get(1);
    	    	JournalName = (String) row.get(2);
    	    	Platform = (String) row.get(5); 
    	    	pptid = (String) row.get(11); 
    		}
    	} catch (Exception e1) {}
    		List<List<Object>> values1 = sheetq.getCellData(sheetsService1, "Details!C3:E3",sheetidq); 
			for (List<?> row : values1) {
				gide = (String) row.get(0);
				gid = (String) row.get(1);
				gidp = (String) row.get(2);
			}
            List<List<Object>> values2 = sheetq.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
            	for (List<?> row : values2) {
            	sheetid = (String) row.get(0);
            }
            	a = 5;
	} catch (Exception e1) {}
    }
        if(!sheetid.contentEquals("")) { 
		ProjGSheet sheet = new ProjGSheet();
		Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
		ProjGSlide pSlide = new ProjGSlide();   		
		Slides service = pSlide.getSlideService("name"); 

		SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("M");
		String datemonth = simpleDateFormat1.format(new Date());  
	
//		System.setProperty("webdriver.edge.driver", "./batchfiles/msedgedriver.exe");
//		HashMap<String, Object> edgePrefs = new HashMap<String, Object>();
//		edgePrefs.put("profile.default_content_settings.popups", 0);
//		edgePrefs.put("download.default_directory", downloadFilepath);
//		EdgeOptions egdeOptions = new EdgeOptions();
//		egdeOptions.setExperimentalOption("prefs",edgePrefs);
//		egdeOptions.addArguments("user-data-dir=C:\\Users\\hake01\\AppData\\Local\\Microsoft\\Edge\\User Data");
//		egdeOptions.addArguments("profile-directory=Work");
//		driver = new EdgeDriver(egdeOptions);
//		driver.get("https://www.google.com/");		
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
//		File fa0 = new File(path1);
//    	if(fa0.exists()){
    		options.addArguments(profile);
//		}else {
//    		options.addArguments("user-data-dir=C:/Users/hake01/AppData/Local/Google/Chrome/User Data/Work");	
//    	} 
    		System.out.println(profile);
    	options.addArguments("--disable-gpu");
    	driver = new ChromeDriver(options);
    	
		String inputN = "";
		int val1=Integer.parseInt(JournalID);  
		int val2 = 10000;

//check
		List<UpdateSlidePojo> updateSlides = new ArrayList<>();	
//		String usage = "No";
//		String usage4 = "No";
//		for (int a = 1; a < 7; a++) {
//		try {
////			driver.get("https://datastudio.google.com/reporting/fce33973-7de7-4812-9d01-6bd7b630d3a6/page/RA5y"); //Usage
//			driver.get("https://datastudio.google.com/reporting/6676196b-cb67-4bc2-94aa-982742414460/page/p_haqd33gfuc"); //rejection
//			driver.manage().window().maximize();
//			Thread.sleep(1000);
//			new WebDriverWait(driver, Duration.ofSeconds(50))
//				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@name='identifier']")));
//				System.out.println(gide);
//				driver.findElement(By.xpath("//*[@id='identifierId']")).sendKeys(gide);
//				Thread.sleep(1000);
//				driver.findElement(By.xpath("(//*[@data-primary-action-label=\"Next\"]//button[1])[1]")).click();
//				new WebDriverWait(driver, Duration.ofSeconds(50))
//				.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='o-form-input']//input)[1]")));
//				driver.findElement(By.xpath("(//*[@class='o-form-input']//input)[1]")).sendKeys(gid); 
//				Thread.sleep(500);
//				driver.findElement(By.xpath("//*[@type='submit']")).click(); 
//
//				new WebDriverWait(driver, Duration.ofSeconds(50))
//				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@type='password']")));
//
//				driver.findElement(By.xpath("//*[@type='password']")).sendKeys(gidp); 
//				Thread.sleep(500);
//				driver.findElement(By.xpath("//*[@type='submit']")).click();
//				try {
//				new WebDriverWait(driver, Duration.ofSeconds(50))
//				.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='VfPpkd-vQzf8d'])[1]")));
//				if(!driver.findElements(By.xpath("(//*[@class='VfPpkd-vQzf8d'])[1]")).isEmpty()){
//				driver.findElement(By.xpath("(//*[@class='VfPpkd-vQzf8d'])[1]")).click();
//				}
//				
//				Thread.sleep(4000);		
//				new WebDriverWait(driver, Duration.ofSeconds(100))
//					      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='dateLabel']")));
//				} catch(Exception e1) {}
//				a = 7;
//				} catch(Exception e1) {}
//		}
		
//rejection tracker	
	for (int a1 = 1; a1 < 5; a1++) {
	String rejectiondata = "";
	File fa01 = new File(downloadFilepath + "Rejection_Tracker.csv");
	File fa01a = new File(downloadFilepath + "MT_2.png");
	File fa01b = new File(downloadFilepath + "MT_2.png");
	if(fa01.exists()){
		rejectiondata = "a";
	}
	if(fa01a.exists()){
		rejectiondata = rejectiondata + "b";
	}
	if(fa01b.exists()){
		rejectiondata = rejectiondata + "c";
	}
	System.out.println(rejectiondata);
	if(!rejectiondata.contentEquals("abc")){
	try {
						driver.get("https://datastudio.google.com/reporting/6676196b-cb67-4bc2-94aa-982742414460/page/p_haqd33gfuc"); //rejection			
						WebElement datelable = new WebDriverWait(driver, Duration.ofSeconds(50))
						      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='dateLabel']")));
						datelable.click();
						WebElement selecttody = new WebDriverWait(driver, Duration.ofSeconds(50))
							      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='md-select-value']")));
							selecttody.click();
							WebElement selecttody1 = new WebDriverWait(driver, Duration.ofSeconds(50))
								      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='_md md-data-studio-theme']//md-option[6]")));
							selecttody1.click();
							WebElement daypicker = new WebDriverWait(driver, Duration.ofSeconds(10))
						      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@Class='uib-daypicker']//button//strong")));
						daypicker.click();
						WebElement monthpicker = new WebDriverWait(driver, Duration.ofSeconds(10))
					      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@Class='uib-monthpicker']//button//strong")));
					monthpicker.click();
					WebElement yearpicker = new WebDriverWait(driver, Duration.ofSeconds(10))
				    .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@Class='uib-yearpicker']//button[1]")));
				yearpicker.click();
				WebElement yeartext = new WebDriverWait(driver, Duration.ofSeconds(10))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='uib-year text-center'])[20]")));
				yeartext.click();
				WebElement uibmonths = new WebDriverWait(driver, Duration.ofSeconds(10))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='uib-months']//td[1]")));
				uibmonths.click();
				WebElement ubiweeks = new WebDriverWait(driver, Duration.ofSeconds(10))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='uib-weeks']//td[4]")));
				ubiweeks.click();
				WebElement buttonbar = new WebDriverWait(driver, Duration.ofSeconds(10))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@Class='button-bar']//button[1]")));
				buttonbar.click();
//						if(!driver.findElements(By.xpath("//*[@Class='uib-daypicker']//button//strong")).isEmpty()){
				Thread.sleep(1000);
						WebElement contentholder = new WebDriverWait(driver, Duration.ofSeconds(50))
						      .until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='content-holder'])[2]")));
						contentholder.click();
						Thread.sleep(5000);
						 
						driver.findElement(By.xpath("(//*[@class='md-label'])[1]")).click();
						WebElement searchb = new WebDriverWait(driver, Duration.ofSeconds(50))
						      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class=\"search-bar\"]//input")));
						searchb.sendKeys(JournalName);
						
						Thread.sleep(5000);
						for (int i1 = 2; i1 < 10; i1++) {
							String valuecheck = driver.findElement(By.xpath("(//*[@class='md-label'])[" + i1 + "]/span")).getText();
							if(valuecheck.contentEquals(JournalName)) {
								driver.findElement(By.xpath("(//*[@class='md-label'])[" + i1 + "]")).click();
								i1 = 10;
							}
						}
						Thread.sleep(4000);
						driver.findElement(By.xpath("//*[@class='editable-label-wrapper']")).click();
						Thread.sleep(7000);
//						driver.findElement(By.xpath("//*[@id='xap-nav-item-4-label']")).click();
//						Thread.sleep(4000);
						if(!fa01.exists()){
						Actions actions1 = new Actions(driver);
						new WebDriverWait(driver, Duration.ofSeconds(200));
						if(!driver.findElements(By.xpath("//*[@title='Published Journal Title']")).isEmpty()){
							WebElement elementLocator1 = driver.findElement(By.xpath("//*[@title='Published Journal Title']"));
							actions1.contextClick(elementLocator1).perform(); 
						}
						Thread.sleep(2000);
						driver.findElement(By.xpath("//div[@id='mat-menu-panel-0']//span[3]/button")).click();
						Thread.sleep(100);
						driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).clear();
						driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).sendKeys("Rejection_Tracker");
						driver.findElement(By.xpath("//div//*[@class='mat-checkbox-inner-container']")).click();                	    
						Thread.sleep(100);
						driver.findElement(By.xpath("//*[@class='mat-dialog-actions gmat-button mat-dialog-actions-align-end']/button[2]")).click(); 
						Thread.sleep(1000); 
						}
						//Screenshots
						driver.findElement(By.xpath("//*[@id='xap-nav-item-3-label']")).click();
						Thread.sleep(5000);
						new WebDriverWait(driver, Duration.ofSeconds(200))
						.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@title='Publisher']")));
						WebElement screen = driver.findElement(By.xpath("(//*[@class=\"content-holder\"])[8]"));
						JavascriptExecutor js = (JavascriptExecutor) driver;
						js.executeScript("arguments[0].scrollIntoView(true);", screen);
						WebElement Re_1 = driver.findElement(By.xpath("(//*[@aria-label=\"A chart.\"])[1]"));
						File fr1 = Re_1.getScreenshotAs(OutputType.FILE);        
						FileUtils.copyFile(fr1, new File(downloadFilepath + "MT_1.png"));
						WebElement screen1 = driver.findElement(By.xpath("(//*[@class='content-holder'])[9]"));
						JavascriptExecutor js1 = (JavascriptExecutor) driver;
						js1.executeScript("arguments[0].scrollIntoView(true);", screen1);
						WebElement Re_2 = driver.findElement(By.xpath("(//*[@aria-label='A chart.'])[3]")); 
						File fr2 = Re_2.getScreenshotAs(OutputType.FILE);
						FileUtils.copyFile(fr2, new File(downloadFilepath + "MT_2.png"));
						Thread.sleep(100);
						a1 = 5;
} catch(Exception e1) {}
}
	}
//Usage
for (int a2 = 1; a2 < 5; a2++) {
String usagedata = "";
File fa02a = new File(downloadFilepath + "4.2.csv");
File fa02b = new File(downloadFilepath + "4.3.csv");
if(fa02a.exists()){
	usagedata = "a";
}
if(fa02b.exists()){
	usagedata = usagedata + "b";
}
System.out.println(usagedata);
if(!usagedata.contentEquals("ab")){
try {
			driver.get("https://datastudio.google.com/reporting/fce33973-7de7-4812-9d01-6bd7b630d3a6/page/RA5y"); 
			Thread.sleep(5000);
			if (val2 <= val1) {
				WebElement bclick = new WebDriverWait(driver, Duration.ofSeconds(50))
						.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='content-holder'])[3]/span[1]")));
				bclick.click();
				inputN = JournalID;
			}else {
				WebElement bclick = new WebDriverWait(driver, Duration.ofSeconds(50))
						.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='content-holder'])[4]/span[1]")));
				bclick.click();
				inputN = JournalName;
			}
			
			Thread.sleep(3000);
			if(!driver.findElements(By.xpath("//*[@aria-label=\"Select all\"]")).isEmpty()){
			driver.findElement(By.xpath("//*[@aria-label=\"Select all\"]")).click();
			Thread.sleep(200);
			driver.findElement(By.xpath("//*[@class='search-bar']//input")).sendKeys(inputN);
			Thread.sleep(1000);
			String patha = "//*[@title=\"" + inputN + "\"]";
			System.out.println(patha);
			driver.findElement(By.xpath(patha)).click(); 
			Thread.sleep(5000);

//			4.2
			Thread.sleep(500);
			driver.findElement(By.id("lego-title-input")).click(); 
			Thread.sleep(500);
			driver.findElement(By.id("xap-nav-item-2-label")).click(); 
			Thread.sleep(10000);
			driver.findElement(By.xpath("(//*[@class=\"content-holder\"])[5]")).click(); 
			Thread.sleep(1000);
			driver.findElement(By.xpath("//*[@title=\"2022\"]")).click();

			Thread.sleep(3000);
			driver.findElement(By.id("lego-title-input")).click(); 
			Thread.sleep(15000);
			Actions actions11 = new Actions(driver);
			try{
				WebElement bclick = new WebDriverWait(driver, Duration.ofSeconds(100))
						.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@aria-label=\"A chart.\"]")));
				actions11.contextClick(bclick).perform(); 
				new WebDriverWait(driver, Duration.ofSeconds(1000))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='mat-menu-panel-0']//span[3]/button")));
			}
			catch(TimeoutException te){
			}
			Thread.sleep(700);
			driver.findElement(By.xpath("//div[@id=\"mat-menu-panel-0\"]//span[3]/button")).click();
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).clear();
			driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).sendKeys("4.2");
//			driver.findElement(By.xpath("(//div//*[@class=\"mat-radio-label\"])[3]")).click();
			driver.findElement(By.xpath("//div//*[@class='mat-checkbox-inner-container']")).click();                	    
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@class='mat-dialog-actions gmat-button mat-dialog-actions-align-end']/button[2]")).click(); 
			Thread.sleep(1000); 
			
//4.3
			Thread.sleep(500);
			driver.findElement(By.id("lego-title-input")).click(); 
			Thread.sleep(500);
			driver.findElement(By.id("xap-nav-item-3-label")).click(); 
			Thread.sleep(10000);
			Actions actionsa = new Actions(driver);
			//try{
			WebElement elet=driver.findElement(By.xpath("//*[@text-anchor=\"start\"]")); 
			JavascriptExecutor jse = (JavascriptExecutor) driver;
			jse.executeScript("arguments[0].scrollIntoView(true);", elet);
			WebElement bclickw = new WebDriverWait(driver, Duration.ofSeconds(1000))
					.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@text-anchor=\"start\"]")));

			WebElement elem1t=driver.findElement(By.xpath("//*[@text-anchor=\"start\"]")); 
			JavascriptExecutor jsq = (JavascriptExecutor) driver;
			jsq.executeScript("arguments[0].scrollIntoView(true);", elem1t);

			actionsa.contextClick(bclickw).perform(); 
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id=\"mat-menu-panel-0\"]//span[3]/button")).click();
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).clear();
			driver.findElement(By.xpath("//*[@class='data-export-dialog']//input")).sendKeys("4.3");
//			driver.findElement(By.xpath("(//div//*[@class=\"mat-radio-label\"])[3]")).click();
			driver.findElement(By.xpath("//div//*[@class='mat-checkbox-inner-container']")).click();                	    
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@class='mat-dialog-actions gmat-button mat-dialog-actions-align-end']/button[2]")).click(); 
			Thread.sleep(5000);
			}	
//4.5
			Thread.sleep(500);
			driver.findElement(By.id("lego-title-input")).click(); 
			Thread.sleep(500);
			driver.findElement(By.id("xap-nav-item-4-label")).click(); 
			Thread.sleep(5000); 
			new WebDriverWait(driver, Duration.ofSeconds(200))
			.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@title='Article Title']")));

int aa = 2;
int ab = 11;
UpdateSlidePojo call1 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call2 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call3 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call4 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call5 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call6 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 4;
ab = 21;
UpdateSlidePojo call11 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call12 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call13 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call14 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call15 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call16 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 5;
ab = 31;
UpdateSlidePojo call21 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call22 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call23 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call24 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call25 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call26 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 7;
ab = 41;
UpdateSlidePojo call31 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call32 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call33 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call34 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call35 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call36 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 8;
ab = 51;
UpdateSlidePojo call41 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call42 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call43 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call44 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call45 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call46 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 9;
ab = 61;
UpdateSlidePojo call51 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call52 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call53 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call54 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call55 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call56 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 10;
ab = 71;
UpdateSlidePojo call61 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call62 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call63 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call64 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call65 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call66 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 11;
ab = 81;
UpdateSlidePojo call71 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call72 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call73 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call74 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call75 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo call76 = new UpdateSlidePojo(46,"<topb" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

WebElement webe=driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[2]")); 
JavascriptExecutor jsaz = (JavascriptExecutor) driver;
jsaz.executeScript("arguments[0].scrollIntoView(true);", webe);

UpdateSlidePojo call7 = new UpdateSlidePojo(46,"<topb17>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[2]")).getText());ab++;
UpdateSlidePojo call17 = new UpdateSlidePojo(46,"<topb27>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[4]")).getText());ab++;
UpdateSlidePojo call27 = new UpdateSlidePojo(46,"<topb37>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[5]")).getText());ab++;
UpdateSlidePojo call37 = new UpdateSlidePojo(46,"<topb47>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[7]")).getText());ab++;
UpdateSlidePojo call47 = new UpdateSlidePojo(46,"<topb57>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[8]")).getText());ab++;
UpdateSlidePojo call57 = new UpdateSlidePojo(46,"<topb67>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[9]")).getText());ab++;
UpdateSlidePojo call67 = new UpdateSlidePojo(46,"<topb77>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[10]")).getText());ab++;
UpdateSlidePojo call77 = new UpdateSlidePojo(46,"<topb87>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[11]")).getText());ab++;

UpdateSlidePojo call8 = new UpdateSlidePojo(46,"<topb18>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[2]")).getText());ab++;
UpdateSlidePojo call18 = new UpdateSlidePojo(46,"<topb28>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[4]")).getText());ab++;
UpdateSlidePojo call28 = new UpdateSlidePojo(46,"<topb38>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[5]")).getText());ab++;
UpdateSlidePojo call38 = new UpdateSlidePojo(46,"<topb48>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[7]")).getText());ab++;
UpdateSlidePojo call48 = new UpdateSlidePojo(46,"<topb58>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[8]")).getText());ab++;
UpdateSlidePojo call58 = new UpdateSlidePojo(46,"<topb68>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[9]")).getText());ab++;
UpdateSlidePojo call68 = new UpdateSlidePojo(46,"<topb78>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[10]")).getText());ab++;
UpdateSlidePojo call78 = new UpdateSlidePojo(46,"<topb88>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[11]")).getText());ab++;

UpdateSlidePojo call9 = new UpdateSlidePojo(46,"<topb19>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[2]")).getText());ab++;
UpdateSlidePojo call19 = new UpdateSlidePojo(46,"<topb29>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[4]")).getText());ab++;
UpdateSlidePojo call29 = new UpdateSlidePojo(46,"<topb39>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[5]")).getText());ab++;
UpdateSlidePojo call39 = new UpdateSlidePojo(46,"<topb49>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[7]")).getText());ab++;
UpdateSlidePojo call49 = new UpdateSlidePojo(46,"<topb59>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[8]")).getText());ab++;
UpdateSlidePojo call59 = new UpdateSlidePojo(46,"<topb69>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[9]")).getText());ab++;
UpdateSlidePojo call69 = new UpdateSlidePojo(46,"<topb79>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[10]")).getText());ab++;
UpdateSlidePojo call79 = new UpdateSlidePojo(46,"<topb89>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[11]")).getText());ab++;

UpdateSlidePojo call10 = new UpdateSlidePojo(46,"<topb20>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[2]")).getText());ab++;
UpdateSlidePojo call20 = new UpdateSlidePojo(46,"<topb30>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[4]")).getText());ab++;
UpdateSlidePojo call30 = new UpdateSlidePojo(46,"<topb40>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[5]")).getText());ab++;
UpdateSlidePojo call40 = new UpdateSlidePojo(46,"<topb50>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[7]")).getText());ab++;
UpdateSlidePojo call50 = new UpdateSlidePojo(46,"<topb60>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[8]")).getText());ab++;
UpdateSlidePojo call60 = new UpdateSlidePojo(46,"<topb70>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[9]")).getText());ab++;
UpdateSlidePojo call70 = new UpdateSlidePojo(46,"<topb80>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[10]")).getText());ab++;
UpdateSlidePojo call80 = new UpdateSlidePojo(46,"<topb90>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[11]")).getText());ab++;

updateSlides.add(call1);
updateSlides.add(call2);
updateSlides.add(call3);
updateSlides.add(call4);
updateSlides.add(call5);
updateSlides.add(call6);
updateSlides.add(call7);
updateSlides.add(call8);
updateSlides.add(call9);
updateSlides.add(call10);
updateSlides.add(call11);
updateSlides.add(call12);
updateSlides.add(call13);
updateSlides.add(call14);
updateSlides.add(call15);
updateSlides.add(call16);
updateSlides.add(call17);
updateSlides.add(call18);
updateSlides.add(call19);
updateSlides.add(call20); 
updateSlides.add(call21);
updateSlides.add(call22);
updateSlides.add(call23);
updateSlides.add(call24);
updateSlides.add(call25);
updateSlides.add(call26);
updateSlides.add(call27);
updateSlides.add(call28);
updateSlides.add(call29);
updateSlides.add(call30); 
updateSlides.add(call31);
updateSlides.add(call32);
updateSlides.add(call33);
updateSlides.add(call34);
updateSlides.add(call35);
updateSlides.add(call36);
updateSlides.add(call37);
updateSlides.add(call38);
updateSlides.add(call39);
updateSlides.add(call40); 
updateSlides.add(call41);
updateSlides.add(call42);
updateSlides.add(call43);
updateSlides.add(call44);
updateSlides.add(call45);
updateSlides.add(call46);
updateSlides.add(call47);
updateSlides.add(call48);
updateSlides.add(call49);
updateSlides.add(call50); 
updateSlides.add(call51);
updateSlides.add(call52);
updateSlides.add(call53);
updateSlides.add(call54);
updateSlides.add(call55);
updateSlides.add(call56);
updateSlides.add(call57);
updateSlides.add(call58);
updateSlides.add(call59);
updateSlides.add(call60); 
updateSlides.add(call61);
updateSlides.add(call62);
updateSlides.add(call63);
updateSlides.add(call64);
updateSlides.add(call65);
updateSlides.add(call66);
updateSlides.add(call67);
updateSlides.add(call68);
updateSlides.add(call69);
updateSlides.add(call70); 
updateSlides.add(call71);
updateSlides.add(call72);
updateSlides.add(call73);
updateSlides.add(call74);
updateSlides.add(call75);
updateSlides.add(call76);
updateSlides.add(call77);
updateSlides.add(call78);
updateSlides.add(call79);
updateSlides.add(call80); 

Thread.sleep(500);
driver.findElement(By.id("lego-title-input")).click(); 
Thread.sleep(500);
driver.findElement(By.id("xap-nav-item-5-label")).click(); 
Thread.sleep(5000); 
new WebDriverWait(driver, Duration.ofSeconds(200))
.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@title='Article Title']")));

aa = 2;
ab = 11;
UpdateSlidePojo callb1 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb2 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb3 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb4 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb5 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb6 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 4;
ab = 21;
UpdateSlidePojo callb11 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb12 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb13 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb14 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb15 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb16 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 5;
ab = 31;
UpdateSlidePojo callb21 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb22 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb23 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb24 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb25 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb26 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 7;
ab = 41;
UpdateSlidePojo callb31 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb32 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb33 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb34 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb35 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb36 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 8;
ab = 51;
UpdateSlidePojo callb41 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb42 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb43 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb44 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb45 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb46 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 9;
ab = 61;
UpdateSlidePojo callb51 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb52 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb53 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb54 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb55 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb56 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 10;
ab = 71;
UpdateSlidePojo callb61 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb62 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb63 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb64 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb65 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb66 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

aa = 11;
ab = 81;
UpdateSlidePojo callb71 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[2]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb72 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[3]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb73 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[4]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb74 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[5]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb75 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[6]//div[" + aa + "]")).getText());ab++;
UpdateSlidePojo callb76 = new UpdateSlidePojo(45,"<topa" + ab + ">",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[7]//div[" + aa + "]")).getText());ab++;

WebElement webb=driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[2]")); 
JavascriptExecutor jsaa = (JavascriptExecutor) driver;
jsaa.executeScript("arguments[0].scrollIntoView(true);", webb);
UpdateSlidePojo callb7 = new UpdateSlidePojo(45,"<topa17>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[2]")).getText());ab++;
UpdateSlidePojo callb17 = new UpdateSlidePojo(45,"<topa27>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[4]")).getText());ab++;
UpdateSlidePojo callb27 = new UpdateSlidePojo(45,"<topa37>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[5]")).getText());ab++;
UpdateSlidePojo callb37 = new UpdateSlidePojo(45,"<topa47>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[7]")).getText());ab++;
UpdateSlidePojo callb47 = new UpdateSlidePojo(45,"<topa57>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[8]")).getText());ab++;
UpdateSlidePojo callb57 = new UpdateSlidePojo(45,"<topa67>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[9]")).getText());ab++;
UpdateSlidePojo callb67 = new UpdateSlidePojo(45,"<topa77>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[10]")).getText());ab++;
UpdateSlidePojo callb77 = new UpdateSlidePojo(45,"<topa87>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[8]//div[11]")).getText());ab++;

UpdateSlidePojo callb8 = new UpdateSlidePojo(45,"<topa18>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[2]")).getText());ab++;
UpdateSlidePojo callb18 = new UpdateSlidePojo(45,"<topa28>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[4]")).getText());ab++;
UpdateSlidePojo callb28 = new UpdateSlidePojo(45,"<topa38>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[5]")).getText());ab++;
UpdateSlidePojo callb38 = new UpdateSlidePojo(45,"<topa48>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[7]")).getText());ab++;
UpdateSlidePojo callb48 = new UpdateSlidePojo(45,"<topa58>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[8]")).getText());ab++;
UpdateSlidePojo callb58 = new UpdateSlidePojo(45,"<topa68>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[9]")).getText());ab++;
UpdateSlidePojo callb68 = new UpdateSlidePojo(45,"<topa78>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[10]")).getText());ab++;
UpdateSlidePojo callb78 = new UpdateSlidePojo(45,"<topa88>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[9]//div[11]")).getText());ab++;

UpdateSlidePojo callb9 = new UpdateSlidePojo(45,"<topa19>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[2]")).getText());ab++;
UpdateSlidePojo callb19 = new UpdateSlidePojo(45,"<topa29>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[4]")).getText());ab++;
UpdateSlidePojo callb29 = new UpdateSlidePojo(45,"<topa39>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[5]")).getText());ab++;
UpdateSlidePojo callb39 = new UpdateSlidePojo(45,"<topa49>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[7]")).getText());ab++;
UpdateSlidePojo callb49 = new UpdateSlidePojo(45,"<topa59>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[8]")).getText());ab++;
UpdateSlidePojo callb59 = new UpdateSlidePojo(45,"<topa69>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[9]")).getText());ab++;
UpdateSlidePojo callb69 = new UpdateSlidePojo(45,"<topa79>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[10]")).getText());ab++;
UpdateSlidePojo callb79 = new UpdateSlidePojo(45,"<topa89>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[10]//div[11]")).getText());ab++;

UpdateSlidePojo callb10 = new UpdateSlidePojo(45,"<topa20>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[2]")).getText());ab++;
UpdateSlidePojo callb20 = new UpdateSlidePojo(45,"<topa30>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[4]")).getText());ab++;
UpdateSlidePojo callb30 = new UpdateSlidePojo(45,"<topa40>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[5]")).getText());ab++;
UpdateSlidePojo callb40 = new UpdateSlidePojo(45,"<topa50>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[7]")).getText());ab++;
UpdateSlidePojo callb50 = new UpdateSlidePojo(45,"<topa60>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[8]")).getText());ab++;
UpdateSlidePojo callb60 = new UpdateSlidePojo(45,"<topa70>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[9]")).getText());ab++;
UpdateSlidePojo callb70 = new UpdateSlidePojo(45,"<topa80>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[10]")).getText());ab++;
UpdateSlidePojo callb80 = new UpdateSlidePojo(45,"<topa90>",driver.findElement(By.xpath("(//*[@class='tableBody'])[2]//div[11]//div[11]")).getText());ab++;

updateSlides.add(callb1);
updateSlides.add(callb2);
updateSlides.add(callb3);
updateSlides.add(callb4);
updateSlides.add(callb5);
updateSlides.add(callb6);
updateSlides.add(callb7);
updateSlides.add(callb8);
updateSlides.add(callb9);
updateSlides.add(callb10);
updateSlides.add(callb11);
updateSlides.add(callb12);
updateSlides.add(callb13);
updateSlides.add(callb14);
updateSlides.add(callb15);
updateSlides.add(callb16);
updateSlides.add(callb17);
updateSlides.add(callb18);
updateSlides.add(callb19);
updateSlides.add(callb20); 
updateSlides.add(callb21);
updateSlides.add(callb22);
updateSlides.add(callb23);
updateSlides.add(callb24);
updateSlides.add(callb25);
updateSlides.add(callb26);
updateSlides.add(callb27);
updateSlides.add(callb28);
updateSlides.add(callb29);
updateSlides.add(callb30); 
updateSlides.add(callb31);
updateSlides.add(callb32);
updateSlides.add(callb33);
updateSlides.add(callb34);
updateSlides.add(callb35);
updateSlides.add(callb36);
updateSlides.add(callb37);
updateSlides.add(callb38);
updateSlides.add(callb39);
updateSlides.add(callb40); 
updateSlides.add(callb41);
updateSlides.add(callb42);
updateSlides.add(callb43);
updateSlides.add(callb44);
updateSlides.add(callb45);
updateSlides.add(callb46);
updateSlides.add(callb47);
updateSlides.add(callb48);
updateSlides.add(callb49);
updateSlides.add(callb50); 
updateSlides.add(callb51);
updateSlides.add(callb52);
updateSlides.add(callb53);
updateSlides.add(callb54);
updateSlides.add(callb55);
updateSlides.add(callb56);
updateSlides.add(callb57);
updateSlides.add(callb58);
updateSlides.add(callb59);
updateSlides.add(callb60); 
updateSlides.add(callb61);
updateSlides.add(callb62);
updateSlides.add(callb63);
updateSlides.add(callb64);
updateSlides.add(callb65);
updateSlides.add(callb66);
updateSlides.add(callb67);
updateSlides.add(callb68);
updateSlides.add(callb69);
updateSlides.add(callb70); 
updateSlides.add(callb71);
updateSlides.add(callb72);
updateSlides.add(callb73);
updateSlides.add(callb74);
updateSlides.add(callb75);
updateSlides.add(callb76);
updateSlides.add(callb77);
updateSlides.add(callb78);
updateSlides.add(callb79);
updateSlides.add(callb80); 

UpdateSlidePojo callz1 = new UpdateSlidePojo(45,"OpenChoice","Open Choice");
UpdateSlidePojo callz2 = new UpdateSlidePojo(45,"FreeAccess","Free Access");
UpdateSlidePojo callz3 = new UpdateSlidePojo(45,"EDITORIALNOTES","Editorial Notes");
UpdateSlidePojo callz4 = new UpdateSlidePojo(45,"REVIEWPAPER","Review Paper");
UpdateSlidePojo callz5 = new UpdateSlidePojo(45,"ORIGINALPAPER","Original Paper");
UpdateSlidePojo callz6 = new UpdateSlidePojo(45,"EVENTS","Events");
UpdateSlidePojo callz7 = new UpdateSlidePojo(45,"LETTER","Letter");

UpdateSlidePojo callzz1 = new UpdateSlidePojo(46,"OpenChoice","Open Choice");
UpdateSlidePojo callzz2 = new UpdateSlidePojo(46,"FreeAccess","Free Access");
UpdateSlidePojo callzz3 = new UpdateSlidePojo(46,"EDITORIALNOTES","Editorial Notes");
UpdateSlidePojo callzz4 = new UpdateSlidePojo(46,"REVIEWPAPER","Review Paper");
UpdateSlidePojo callzz5 = new UpdateSlidePojo(46,"ORIGINALPAPER","Original Paper");
UpdateSlidePojo callzz6 = new UpdateSlidePojo(46,"EVENTS","Events");
UpdateSlidePojo callzz7 = new UpdateSlidePojo(46,"LETTER","Letter");

updateSlides.add(callz1); 
updateSlides.add(callz2); 
updateSlides.add(callz3); 
updateSlides.add(callz4); 
updateSlides.add(callz5); 
updateSlides.add(callz6); 
updateSlides.add(callz7); 
updateSlides.add(callzz1); 
updateSlides.add(callzz2); 
updateSlides.add(callzz3); 
updateSlides.add(callzz4); 
updateSlides.add(callzz5); 
updateSlides.add(callzz6); 
updateSlides.add(callzz7); 
a2=5;
} catch(Exception e1) {}
}
}
//Usage 4
for (int a3 = 1; a3 < 5; a3++) {
File fa03 = new File(downloadFilepath + "Country.csv");
if(!fa03.exists()){
try {
if(Platform.contentEquals("Springer")) {
	driver.get("https://datastudio.google.com/reporting/1DP10VbpwIv3dCXvqk0D-LjhshIeaa9Oo/page/taj8?pli=1");
}
if(Platform.contentEquals ("BMC")) {
	driver.get("https://datastudio.google.com/reporting/1sWFgUf71UOEgSTUG9tTUu0Y8nP652kr8/page/Evi9?pli=1");
}
if(Platform.contentEquals("Nature")) {
	driver.get("https://datastudio.google.com/u/0/reporting/1BQ0fu8ct8bV0L5lTU7Tl_CB5tqu921M8/page/nYl8");
}
				//log
				Thread.sleep(10000);
	    	    new WebDriverWait(driver, Duration.ofSeconds(50))
	    		.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class=\"datepicker\"]")));
	    	    if(!driver.findElements(By.xpath("//*[@class='datepicker']")).isEmpty()){
	    	    driver.findElement(By.xpath("//*[@class=\"datepicker\"]")).click(); 
	    	    driver.findElement(By.xpath("//*[@Class='md-select-value']")).click();	  
	    	    Thread.sleep(500);
	    	    driver.findElement(By.xpath("//*[@Class='_md md-data-studio-theme']//md-option[16]")).click();
	    	    driver.findElement(By.xpath("//*[@Class='uib-daypicker']//button//strong")).click();
	    	    driver.findElement(By.xpath("//*[@Class='uib-monthpicker']//button//strong")).click();
	    	    driver.findElement(By.xpath("//*[@class='uib-yearpicker']//tr[1]//td[1]")).click();
	    	    
	    	    int monthinnumber = Integer.parseInt(datemonth);
	    	    int monthrow = 1;  
	    	    int getmonth = 1;
	    	    	if (monthinnumber < 4) {
	    	    		monthrow = 1;
	    	    	}
	    	    	if (monthinnumber > 3) {
	    	    		if (monthinnumber < 7) {
	    	    		monthrow = 2;
	    	    		}
	    	    	}
	    	    	if (monthinnumber > 6) {
	    	    		if (monthinnumber < 10) {
	    	    		monthrow = 3;
	    	    		}
	    	    	}
	    	    	if (monthinnumber > 9) {
	    	    		if (monthinnumber < 13) {
	    	    		monthrow = 4;
	    	    		}
	    	    	}	    	    	
	    	    	if (monthinnumber == 1) {getmonth = 1;}
	    	    	if (monthinnumber == 2) {getmonth = 2;}
	    	    	if (monthinnumber == 3) {getmonth = 3;}
	    	    	if (monthinnumber == 4) {getmonth = 1;}
	    	    	if (monthinnumber == 5) {getmonth = 2;}
	    	    	if (monthinnumber == 6) {getmonth = 3;}
	    	    	if (monthinnumber == 7) {getmonth = 1;}
	    	    	if (monthinnumber == 8) {getmonth = 2;}
	    	    	if (monthinnumber == 9) {getmonth = 3;}
	    	    	if (monthinnumber == 10) {getmonth = 1;}
	    	    	if (monthinnumber == 11) {getmonth = 2;}
	    	    	if (monthinnumber == 12) {getmonth = 3;}
	    	    	driver.findElement(By.xpath("//*[@class='uib-monthpicker']//tr[" + monthrow + "]//td[" + getmonth + "]")).click();
	    	    	for (int checkcount = 1; checkcount < 8; checkcount++) {
	    	    	try {
	    	    driver.findElement(By.xpath("((//*[@class=\"uib-daypicker\"])[1]//tr[1]//*[@class=\"text-muted\"])[" + checkcount + "]")).getText();
	    	    	} catch(Exception e1) {
	    	    		try {
	    	    		driver.findElement(By.xpath("(//*[@class='uib-daypicker']//tr[1]//td[" + checkcount + "])[1]")).click();
	    	    	    driver.findElement(By.xpath("//*[@Class='button-bar']//button[1]")).click();
	    	    		} catch(Exception e2) {}
	    	    }
		}
//check
	    	    	Thread.sleep(7000);
					if(Platform.contentEquals("Nature")) {
						WebElement bclick = new WebDriverWait(driver, Duration.ofSeconds(50))
								.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='content-holder'])[2]")));
						bclick.click();
						Thread.sleep(500);
						driver.findElement(By.xpath("//*[@aria-label='Select all']")).click();
						Thread.sleep(500);					
						String lowercaseJournalName = JournalName.toLowerCase(); 
						driver.findElement(By.xpath("//*[@class=\"search-bar\"]//input")).sendKeys(lowercaseJournalName);
						Thread.sleep(2000);					
//						driver.findElement(By.xpath("//*[@class='metric-value']")).click();					
						if(!driver.findElements(By.xpath("//*[@title='" + lowercaseJournalName + "']")).isEmpty()){
							driver.findElement(By.xpath("//*[@title='" + lowercaseJournalName + "']")).click();	
						}else {
							driver.findElement(By.xpath("//*[@class=\"search-bar\"]//input")).sendKeys(JournalName);
							try {
							driver.findElement(By.xpath("//*[@title='" + JournalName + "']")).click();	
						} catch(Exception e1) {driver.close();}
						}
					}else {
						Thread.sleep(100);
//						if (val2 <= val1) {
							WebElement bclick = new WebDriverWait(driver, Duration.ofSeconds(50))
									.until(ExpectedConditions.elementToBeClickable(By.xpath("(//*[@class='content-holder'])[1]")));
							bclick.click();
						Thread.sleep(1000);
						driver.findElement(By.xpath("//*[@aria-label=\"Select all\"]")).click();
						driver.findElement(By.xpath("//*[@class=\"search-bar\"]//input")).sendKeys(JournalID); 
						Thread.sleep(1000);
						try {
						driver.findElement(By.xpath("//*[@title='" + JournalID + "']")).click();
						} catch(Exception e1) {driver.close();}
					}
					driver.findElement(By.xpath("//*[@id='lego-title-input']")).click();
	    	    	
//check		
				Thread.sleep(10000);
				driver.findElement(By.xpath("//*[@id='lego-title-input']")).click();
				new WebDriverWait(driver, Duration.ofSeconds(20))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='xap-nav-item-1-label']")));
				
				driver.findElement(By.id("xap-nav-item-1-label")).click(); 
				Thread.sleep(5000);
				new WebDriverWait(driver, Duration.ofSeconds(50))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@title='Traffic Source']")));
				String total = driver.findElement(By.xpath("(//*[@class='totalsContent']//span)[3]")).getText();
				String cala1 = "";
				String calb2 = ""; 
				String middalvalue = "";
				int i2 = 3;
				for (int i1 = 2; i1 < 7; i1++) {
				String vala = driver.findElement(By.xpath("//*[@class='tableBody']/div[" + i1 + "]/div[2]")).getText();
				String valb = driver.findElement(By.xpath("//*[@class='tableBody']/div[" + i1 + "]/div[3]")).getText();
				cala1 = cala1 + middalvalue + vala;
				calb2 = calb2 + middalvalue + valb;
				middalvalue = "^";
				i2 = i2 + 1;
				}
				sheet.updateCellvalue(sheetsService, sheetname + "!A14", cala1,sheetid);
				sheet.updateCellvalue(sheetsService, sheetname + "!A15", calb2,sheetid);
				sheet.updateCellvalue(sheetsService, sheetname + "!A19", total,sheetid);
				
				Actions actions = new Actions(driver);
				driver.findElement(By.id("xap-nav-item-2-label")).click(); 
				Thread.sleep(100);

				driver.findElement(By.id("xap-nav-item-2-label")).click(); 
				Thread.sleep(10000);

				new WebDriverWait(driver, Duration.ofSeconds(50))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@title='Country']")));
				if(!driver.findElements(By.xpath("//*[@title='Country']")).isEmpty()){
					WebElement elementLocator1 = driver.findElement(By.xpath("//*[@title='Country']"));
					actions.contextClick(elementLocator1).perform(); 
				}
				Thread.sleep(500);
				new WebDriverWait(driver, Duration.ofSeconds(50))
				.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='mat-menu-panel-0']/div/span[3]/button")));
				driver.findElement(By.xpath("//*[@id='mat-menu-panel-0']/div/span[3]/button")).click();
				Thread.sleep(100);
				driver.findElement(By.xpath("//*[@aria-label=\"Choose the export file type\"]//mat-radio-button[3]")).click();
				driver.findElement(By.xpath("//*[@class=\"mat-checkbox-inner-container\"]")).click();
				driver.findElement(By.xpath("//*[@class=\"data-export-dialog\"]//input")).clear();
				driver.findElement(By.xpath("//*[@class=\"data-export-dialog\"]//input")).sendKeys("Country");
				driver.findElement(By.xpath("(//div//*[@class=\"mat-radio-label\"])[2]")).click();
				driver.findElement(By.xpath("//*[@class=\"mat-dialog-actions gmat-button mat-dialog-actions-align-end\"]/button[2]")).click(); 
				Thread.sleep(5000);
				}
//usage4 = "Yes";
a3 = 5;
} catch(Exception e1) {System.out.println("IOException");}
}
try {
pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides);	
} catch(Exception e1) {}
}
try {
			driver.close();
			driver.quit();
	} catch(Exception e1) {System.out.println("IOException");}
}
}
//}
  @AfterTest
  public void afterTest() {	 
	  try {
			driver.close();
			driver.quit();
		} catch(Exception e1) {System.out.println("Exceptionbex");}
  }

}