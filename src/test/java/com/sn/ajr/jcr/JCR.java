package com.sn.ajr.jcr;
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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions; 
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
public class JCR {
	public static WebDriver driver;
  @Test( timeOut = 400000 )
  public static void classname() throws IOException, GeneralSecurityException, InterruptedException { 
	  System.out.println("JCR start");
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
	  	String pISSN = "";
	  	String eISSN = ""; 
	  	String sheetname = "";
	  	String pptid = "";
	  	String wosid = "";
	  	String wosp = "";
	  	try {
	  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
			for (List<?> row : values) {
		sheetname = (String) row.get(0);
	  	pISSN = (String) row.get(3);
	  	eISSN = (String) row.get(4);	  	
		pptid = (String) row.get(11); 
			}
			List<List<Object>> values1 = sheetq.getCellData(sheetsService1, "Details!C5:E5",sheetidq); 
			for (List<?> row : values1) {
				wosid = (String) row.get(0);
				wosp = (String) row.get(2);
			}
	  	} catch (Exception e1) {}
		
		if (eISSN.contentEquals("")){
			eISSN = pISSN;
		}
		
	  	if(!eISSN.contentEquals("")) {
	  		String sheetid = "";
        	try {
        	List<List<Object>> values2 = sheetq.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
        	for (List<?> row : values2) {
        	sheetid = (String) row.get(0);
        	}
        	 } catch (Exception e1) {}
        	if(!sheetid.contentEquals("")) { 	
		ProjGSheet sheet = new ProjGSheet();
		Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
		
		ProjGSlide pSlide = new ProjGSlide();
		Slides service = pSlide.getSlideService("name"); 
		
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--disable-gpu");
		driver = new ChromeDriver(options); 
  	
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM");
		String date1 = simpleDateFormat.format(new Date());
		 System.out.println("check" + wosid);
try {
//web of science clarivate
	for (int a = 1; a < 5; a++) {
      driver.get("https://access.clarivate.com/login?app=jcr&referrer=target%3Dhttps:%2F%2Fjcr.clarivate.com%2Fjcr%2Fhom");
      driver.manage().window().maximize();

      Thread.sleep(5000);
      if(!driver.findElements(By.xpath("//*[@aria-label='link button']/span")).isEmpty()){
          driver.findElement(By.xpath("//*[@aria-label='link button']/span")).click();}
      
      if(!driver.findElements(By.xpath("//*[@name='email']")).isEmpty()){ 
          driver.findElement(By.xpath("//*[@name='email']")).sendKeys(wosid);
          driver.findElement(By.xpath("//*[@name='password']")).sendKeys(wosp);
          driver.findElement(By.id("signIn-btn")).click();
      }
      Thread.sleep(7000); 
      if(!driver.findElements(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).isEmpty()){
      	driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).click();
      }
      if(!driver.findElements(By.className("_pendo-close-guide")).isEmpty()){ 
      	driver.findElement(By.className("_pendo-close-guide")).click();
      }
      WebElement searchlable = new WebDriverWait(driver, Duration.ofSeconds(10))
      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='search-bar']")));
      searchlable.sendKeys(eISSN);
      Thread.sleep(7000);
      
      if(!driver.findElements(By.xpath("//*[@class='pop-content journal-title']")).isEmpty()){
      WebElement search1 = new WebDriverWait(driver, Duration.ofSeconds(10))
      			.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='pop-content journal-title']")));
      	try{
      		if(!driver.findElements(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).isEmpty()){
              	driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).click();
              }
      	} catch(Exception e1) {}
      	search1.click();
      	Thread.sleep(10000); 
      	@SuppressWarnings("unused")
			String winHandleBefore = driver.getWindowHandle();
      	for(String winHandle : driver.getWindowHandles()){
      		driver.switchTo().window(winHandle);
      	}
      	Thread.sleep(10000); 
      	if(!driver.findElements(By.xpath("//*[@class='calculation']")).isEmpty()){
      	Thread.sleep(3000); 
      	String firstyearvalue = ""; 
      	String secondyearvalue = ""; 
      	String thirdyearvalue = ""; 
      	String fouryearvalue = ""; 
      	String fiveyearvalue = ""; 
      	WebElement yearelement1 = driver.findElement(By.xpath("//*[@class='calculation']")); 
      	JavascriptExecutor jsa1 = (JavascriptExecutor) driver;
      	WebElement valueelement2 = driver.findElement(By.xpath("//*[@role='combobox']")); 
      	JavascriptExecutor jsa2 = (JavascriptExecutor) driver;
      	try {
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2017");
      	Thread.sleep(2000);
      	firstyearvalue = "0" + "^" + "0"; 
      	if (driver.findElement(By.xpath("//*[@role='combobox']")).getText().contentEquals("2017")) {
      		jsa1.executeScript("arguments[0].scrollIntoView(true);", yearelement1);	
      		Thread.sleep(300);
      		driver.findElement(By.xpath("//*[@class='calculation']")).click();
      		Thread.sleep(200);
      		String val1 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      		String val2 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[2]")).getText();
      		driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();
      		firstyearvalue = val1 + "^" + val2; 
      	}
      	} catch(Exception e1) {}
      	try {
      	jsa1.executeScript("arguments[0].scrollIntoView(true);", valueelement2);
      	Thread.sleep(200);
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2018");
      	Thread.sleep(2000);
      	secondyearvalue = "0" + "^" + "0"; 
      	if (driver.findElement(By.xpath("//*[@role='combobox']")).getText().contentEquals("2018")) {
      		jsa2.executeScript("arguments[0].scrollIntoView(true);", yearelement1);	
      		Thread.sleep(500);
      		driver.findElement(By.xpath("//*[@class='calculation']")).click();
      		Thread.sleep(200);
      		String val1 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      		String val2 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[2]")).getText();
      		driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();
      		secondyearvalue = val1 + "^" + val2; 
      	}
      	} catch(Exception e1) {}
      	try {
      	jsa1.executeScript("arguments[0].scrollIntoView(true);", valueelement2);
      	Thread.sleep(200);
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2019");
      	Thread.sleep(2000);
      	thirdyearvalue = "0" + "^" + "0";
      	if (driver.findElement(By.xpath("//*[@role='combobox']")).getText().contentEquals("2019")) {
      		jsa2.executeScript("arguments[0].scrollIntoView(true);", yearelement1);	
      		Thread.sleep(100);
      		driver.findElement(By.xpath("//*[@class='calculation']")).click();
      		Thread.sleep(200);
      		String val1 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      		String val2 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[2]")).getText();
      		driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();
      		thirdyearvalue = val1 + "^" + val2; 
      	}
      	} catch(Exception e1) {}
      	try {
      	jsa1.executeScript("arguments[0].scrollIntoView(true);", valueelement2);
      	Thread.sleep(200);
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2020");
      	Thread.sleep(2000);
      	fouryearvalue = "0" + "^" + "0"; 
      	if (driver.findElement(By.xpath("//*[@role='combobox']")).getText().contentEquals("2020")) {
      		jsa2.executeScript("arguments[0].scrollIntoView(true);", yearelement1);	
      		Thread.sleep(300);
      		driver.findElement(By.xpath("//*[@class='calculation']")).click();
      		Thread.sleep(200);
      		String val1 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      		String val2 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[2]")).getText();
      		driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();
      		fouryearvalue = val1 + "^" + val2; 
      	}
      	} catch(Exception e1) {}
      	try {
      	jsa1.executeScript("arguments[0].scrollIntoView(true);", valueelement2);
      	Thread.sleep(200);
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2021");
      	Thread.sleep(2000);
      	fiveyearvalue = "0" + "^" + "0";
      	if (driver.findElement(By.xpath("//*[@role='combobox']")).getText().contentEquals("2021")) {
      		jsa2.executeScript("arguments[0].scrollIntoView(true);", yearelement1);	
      		Thread.sleep(300);
      		driver.findElement(By.xpath("//*[@class='calculation']")).click();
      		Thread.sleep(200);
      		String val1 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      		String val2 = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[2]")).getText();
      		driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();
      		fiveyearvalue = val1 + "^" + val2; 
      	}
      	} catch(Exception e1) {}

      	
      	try {
      	jsa1.executeScript("arguments[0].scrollIntoView(true);", valueelement2);
      	Thread.sleep(200);
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2021");
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2022");
      	} catch(Exception e1) {}
      	
      	String	catm1 = "";
      	String	catm2 = "";
      	String	catm3 = "";
      	String	catm4 = "";
      	String	catm5 = "";
      	String	catm6 = "";
      	String	catm7 = "";
      	
      	String cat1 = "";
      	String cat1r = "";
      	String cat1q = "";
      	
      	String cat2 = "";
      	String cat2r = "";
      	String cat2q = "";
      	
      	String cat3 = "";
      	String cat3r = "";
      	String cat3q = "";

      	String cat4 = "";
      	String cat4r = "";
      	String cat4q = "";
      	
      	String cat5 = "";
      	String cat5r = "";
      	String cat5q = "";
      	
      	String cat6 = "";
      	String cat6r = "";
      	String cat6q = "";
      	
      	String cat7 = "";
      	String cat7r = "";
      	String cat7q = "";
      		
//      		1
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-1']//div[4]")).isEmpty()){
      		cat1 = driver.findElement(By.xpath("//*[@id='rbj-silde-1']//div[4]")).getText();
      		cat1r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[1]//tr[2]//td[2]")).getText();
      		cat1q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[1]//tr[2]//td[3]")).getText(); 
      		catm1 = cat1 + '^' + cat1r + '^' + cat1q;  
      		}
//      		2
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-2']//div[4]")).isEmpty()){
      		cat2 = driver.findElement(By.xpath("//*[@id='rbj-silde-2']//div[4]")).getText();
      		cat2r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[2]//tr[2]//td[2]")).getText();
      		cat2q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[2]//tr[2]//td[3]")).getText(); 
      		catm2 = cat2 + '^' + cat2r + '^' + cat2q;
      		}
//      		3
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-3']//div[4]")).isEmpty()){
      		cat3 = driver.findElement(By.xpath("//*[@id='rbj-silde-3']//div[4]")).getText();
      		cat3r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[3]//tr[2]//td[2]")).getText();
      		cat3q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[3]//tr[2]//td[3]")).getText(); 
      		catm3 = cat3 + '^' + cat3r + '^' + cat3q;
      		}
//      		4
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-4']//div[4]")).isEmpty()){
      		cat4 = driver.findElement(By.xpath("//*[@id='rbj-silde-4']//div[4]")).getText();
      		cat4r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[4]//tr[2]//td[2]")).getText();
      		cat4q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[4]//tr[2]//td[3]")).getText(); 
      		catm4 = cat4 + '^' + cat4r + '^' + cat4q;
      		}
//      		5
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-5']//div[4]")).isEmpty()){
      		cat5 = driver.findElement(By.xpath("//*[@id='rbj-silde-5']//div[4]")).getText();
      		cat5r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[5]//tr[2]//td[2]")).getText();
      		cat5q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[5]//tr[2]//td[3]")).getText(); 
      		catm5 = cat5 + '^' + cat5r + '^' + cat5q;
      		}
//      		6
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-6']//div[4]")).isEmpty()){
      		cat6 = driver.findElement(By.xpath("//*[@id='rbj-silde-6']//div[4]")).getText();
      		cat6r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[6]//tr[2]//td[2]")).getText();
      		cat6q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[6]//tr[2]//td[3]")).getText();
      		catm6 = cat6 + '^' + cat6r + '^' + cat6q;
      		}
//      		7
      		if(!driver.findElements(By.xpath("//*[@id='rbj-silde-7']//div[4]")).isEmpty()){
      		cat7 = driver.findElement(By.xpath("//*[@id='rbj-silde-7']//div[4]")).getText();
      		cat7r = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[7]//tr[2]//td[2]")).getText();
      		cat7q = driver.findElement(By.xpath("(//*[@aria-describedby='Rank by Journal Impact Factor'])[7]//tr[2]//td[3]")).getText(); 
      		catm7 = cat7 + '^' + cat7r + '^' + cat7q;
      		}
      		
          	String catmmm = catm1 + '\n' + catm2 + '\n' + catm3 + '\n' + catm4 + '\n' + catm5 + '\n' + catm6 + '\n' + catm7;
          	String	catmm = catmmm.replace('/', '^');
//          	System.out.println(catmm);
      		sheet.updateCellvalue(sheetsService, sheetname + "!R159", catmm ,sheetid);
      //Categorys
     try {
      //Cat1
      	WebElement catelementz=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[1]")); 
      	JavascriptExecutor jsaz = (JavascriptExecutor) driver;
      	jsaz.executeScript("arguments[0].scrollIntoView(true);", catelementz);		
      	Thread.sleep(2000);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[1]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[1]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S2", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[1]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(3000);
              if(!driver.findElements(By.xpath("//*[@title='Download']")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title='Download']")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();  
          
      		File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
      		if(myfile.exists()){
      			myfile.renameTo(new File(downloadFilepath + "\\JCR_category1.xlsx"));	
      		}else {
      			date1 = simpleDateFormat.format(new Date());
          		myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
          		if(myfile.exists()){
          			myfile.renameTo(new File(downloadFilepath + "\\JCR_category1.xlsx"));
          		}
      		}
      		
      		
      		
          }
//          cat1 end
      } catch(Exception e1) {}
       
          try {
      	//Cat2
      	WebElement catelementq=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[2]")); 
      	JavascriptExecutor jsq = (JavascriptExecutor) driver;
      	jsq.executeScript("arguments[0].scrollIntoView(true);", catelementq);	
      	Thread.sleep(1000);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[2]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[2]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S28", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[2]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(1000);
              if(!driver.findElements(By.xpath("//*[@title=\"Download\"]")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title=\"Download\"]")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();  
            
              File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
              if(myfile.exists()){
            	  myfile.renameTo(new File(downloadFilepath + "\\JCR_category2.xlsx"));
        		}else {
        			date1 = simpleDateFormat.format(new Date());
        			myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
        			myfile.renameTo(new File(downloadFilepath + "\\JCR_category2.xlsx"));
        		}
      		
      		
      	}
          } catch(Exception e1) {}
//  	    cat2 end
          
      try {
      	//Cat3
      	WebElement catelement=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[3]")); 
      	JavascriptExecutor jsa = (JavascriptExecutor) driver;
      	jsa.executeScript("arguments[0].scrollIntoView(true);", catelement);	
      	Thread.sleep(1000);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[3]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[3]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S54", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[3]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(1000);
              if(!driver.findElements(By.xpath("//*[@title=\"Download\"]")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title=\"Download\"]")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();
              File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
              if(myfile.exists()){
            	  myfile.renameTo(new File(downloadFilepath + "\\JCR_category3.xlsx"));
        		}else {
        			date1 = simpleDateFormat.format(new Date());
        			myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
        			myfile.renameTo(new File(downloadFilepath + "\\JCR_category3.xlsx"));
        		}
          }
      	} catch(Exception e1) {}
      	    
      try {
      	//Cat4
      	WebElement catelement=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[4]")); 
      	JavascriptExecutor jsa = (JavascriptExecutor) driver;
      	jsa.executeScript("arguments[0].scrollIntoView(true);", catelement);		
      	Thread.sleep(500);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[4]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[4]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S80", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[4]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(500);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(1000);
              if(!driver.findElements(By.xpath("//*[@title=\"Download\"]")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title=\"Download\"]")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();  
              File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
              if(myfile.exists()){
            	  myfile.renameTo(new File(downloadFilepath + "\\JCR_category4.xlsx"));
        		}else {
        			date1 = simpleDateFormat.format(new Date());
        			myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
        			myfile.renameTo(new File(downloadFilepath + "\\JCR_category4.xlsx"));
        		}
          }

      	} catch(Exception e1) {}

      try {
      	//Cat5
      	WebElement catelement=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[5]")); 
      	JavascriptExecutor jsa = (JavascriptExecutor) driver;
      	jsa.executeScript("arguments[0].scrollIntoView(true);", catelement);	
      	Thread.sleep(1000);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[5]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[5]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S106", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[5]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(500);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(1000);
              if(!driver.findElements(By.xpath("//*[@title=\"Download\"]")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title=\"Download\"]")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();  
              File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
              if(myfile.exists()){
            	  myfile.renameTo(new File(downloadFilepath + "\\JCR_category5.xlsx"));
        		}else {
        			date1 = simpleDateFormat.format(new Date());
        			myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
        			myfile.renameTo(new File(downloadFilepath + "\\JCR_category5.xlsx"));
        		}
          }
      	} catch(Exception e1) {}

      try {
      	//Cat6
      	WebElement catelement=driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[6]")); 
      	JavascriptExecutor jsa = (JavascriptExecutor) driver;
      	jsa.executeScript("arguments[0].scrollIntoView(true);", catelement);	
      	Thread.sleep(1000);
          if(!driver.findElements(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[6]")).isEmpty()){
          	String categoryname1 =  driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[6]")).getText();
          	sheet.updateCellvalue(sheetsService, sheetname + "!S132", categoryname1,sheetid);
          	driver.findElement(By.xpath("(//*[@class='category-text info-tile-p-text ng-star-inserted'])[6]")).click();
          	Thread.sleep(2000);    	
          	driver.findElement(By.xpath("//*[@class='settings-icon']")).click();
          	Thread.sleep(500);
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[4]//span")).click();
          	driver.findElement(By.xpath("//*[@class='row indicator-options']//mat-checkbox[5]//span")).click();
          	driver.findElement(By.xpath("(//*[@aria-label='Apply'])[2]//span[1]")).click();    	
          	Thread.sleep(1000);
          	driver.findElement(By.xpath("(//*[@class=\"table-section\"]//mat-header-cell)[6]")).click();
          	Thread.sleep(1000);
              if(!driver.findElements(By.xpath("//*[@title=\"Download\"]")).isEmpty()){ 
              	driver.findElement(By.xpath("//*[@title=\"Download\"]")).click();
              	Thread.sleep(100);
              	driver.findElement(By.xpath("(//*[@role='menu']//button)[2]")).click();
              }
              Thread.sleep(5000);
              driver.navigate().back();  
              File myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
              if(myfile.exists()){
            	  myfile.renameTo(new File(downloadFilepath + "\\JCR_category6.xlsx"));
        		}else {
        			date1 = simpleDateFormat.format(new Date());
        			myfile = new File(downloadFilepath + "\\ManasiKapne_JCR_JournalResults_" + date1 + "_2022.xlsx");
        			myfile.renameTo(new File(downloadFilepath + "\\JCR_category6.xlsx"));
        		}
          }
} catch(Exception e1) {}

if(!driver.findElements(By.xpath("//*[@id=\"onetrust--btn-handler\"]")).isEmpty()){
      			driver.findElement(By.xpath("//*[@id=\"onetrust--btn-handler\"]")).click();
     }
//System.out.println("start persent check");
      try {
//      	Cited and citing
      new WebDriverWait(driver, Duration.ofSeconds(1000))
      			.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='incites-jcr3-fe-export row text-right ng-star-inserted']//mat-icon")));
      Thread.sleep(10000);
      	if(!driver.findElements(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).isEmpty()){
      		driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]")).click();
      	}
      	
      	driver.findElement(By.xpath("//*[@role='combobox']")).sendKeys("2021");
      	WebElement element=driver.findElement(By.xpath("(//*[@id=\"mat-tab-content-0-0\"]/div/div/div)[1]")); 
      	JavascriptExecutor js = (JavascriptExecutor) driver;
      	js.executeScript("arguments[0].scrollIntoView(true);", element);
      	Thread.sleep(1000);
      	js.executeScript("arguments[0].scrollIntoView(true);", element);
//      	driver.findElement(By.xpath("(//*[@id=\"mat-tab-content-0-0\"]/div/div/div)[1]")).click();
      	driver.findElement(By.xpath("//*[@class='incites-jcr3-fe-export row text-right ng-star-inserted']//span")).click();
      	Thread.sleep(10000);
      	driver.findElement(By.xpath("//*[@aria-label='Export XLS']//mat-icon")).click();
      	Thread.sleep(2500);
      	driver.findElement(By.xpath("(//*[@Class='half-life-section']//div)[8]")).click();
      	Thread.sleep(5000);
      	driver.findElement(By.xpath("//*[@class='incites-jcr3-fe-export row text-right ng-star-inserted']//mat-icon")).click();
      	Thread.sleep(500);
      	driver.findElement(By.xpath("//*[@aria-label='Export XLS']//mat-icon")).click();
      	Thread.sleep(10000);
      	 } catch (Exception ex) {}

      try {
      Thread.sleep(3000);
      //otherinfo
      	WebElement eleme = driver.findElement(By.xpath("(//*[@class=\"incites-jcr3-fe-jif-trend\"]//p[3])[2]")); 
      	JavascriptExecutor jsy = (JavascriptExecutor) driver;
      	jsy.executeScript("arguments[0].scrollIntoView(true);", eleme);
      	driver.findElement(By.xpath("(//*[@class=\"incites-jcr3-fe-jif-trend\"]//p[3])[2]")).click();
			Thread.sleep(200);
      	String withocal = driver.findElement(By.xpath("//*[@class='calculation-columns container3']//p[1]")).getText();
      	
      	driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click(); 
      	Thread.sleep(5000);
      	try {
         	driver.findElement(By.xpath("//*[@class='incites-jcr3-fe-five-year-impact-factor']//p[3]")).click();
      	WebElement SJRlogo1 = driver.findElement(By.xpath("//*[@class='mat-dialog-content']"));
      	Thread.sleep(500);
         	File fs2 = SJRlogo1.getScreenshotAs(OutputType.FILE);    
         	try {
				FileUtils.copyFile(fs2, new File(downloadFilepath + "\\5YearIF.png"));
			} catch (IOException e1) {}
         	driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click();   	
         	Thread.sleep(200);
      	} catch(Exception e1) {}
      	
         	String IFwith = driver.findElement(By.xpath("(//*[@class='incites-jcr3-fe-jif-trend']//div[1]//p[2])[1]")).getText();
         	String IFwithout = driver.findElement(By.xpath("(//*[@class='incites-jcr3-fe-jif-trend']//p[2])[2]")).getText();
         	String fiveyear = driver.findElement(By.xpath("//*[@class='five-yr-impact-factor-value']")).getText();
         	
         	driver.findElement(By.xpath("(//*[@class='incites-jcr3-fe-jif-trend']//p[3])[1]")).click();
      	WebElement SJRlogo = driver.findElement(By.xpath("//*[@class='mat-dialog-content']"));
      	Thread.sleep(500);
         	File fs1 = SJRlogo.getScreenshotAs(OutputType.FILE);        
         	try {
				FileUtils.copyFile(fs1, new File(downloadFilepath + "\\impact-factor.png"));
			} catch (IOException e1) {e1.printStackTrace();}
         	Thread.sleep(1000);
         	driver.findElement(By.xpath("(//*[@aria-label='Close'])[2]")).click(); 	
      	String part1 = "";
      	String part2 = "";
      	double persent = 0;
      	try {
      	try {
      		withocal = withocal.replace(",", "");
      	} catch (Exception ex) {}
      	String[] parts0 = withocal.split(" - ");
			part1 = parts0[0];
			part2 = parts0[1];
//			System.out.println(withocal);
//			System.out.println(parts0);
//			System.out.println(part1);
      	 double withocalvalue = Integer.parseInt(part1);
 		 double IFwithvalue = Integer.parseInt(part2);
 		 persent = (IFwithvalue / withocalvalue) * 100; 
// 		 System.out.println("persent check" + withocalvalue + "&" + IFwithvalue + "=" + persent);
      	} catch (Exception ex) {}
       try {
       sheet.updateCellvalue(sheetsService, sheetname + "!H145", "AImp_factorA" + "|" + firstyearvalue + "|" + secondyearvalue + "|" + thirdyearvalue + "|" + fouryearvalue + "|" + fiveyearvalue,sheetid);
      } catch (Exception ex) {}
//System.out.println("AImp_factorA" + "|" + firstyearvalue + "|" + secondyearvalue + "|" + thirdyearvalue + "|" + fouryearvalue + "|" + fiveyearvalue);
      	
      	List<UpdateSlidePojo> updateSlides = new ArrayList<>();
          UpdateSlidePojo call1 = new UpdateSlidePojo(56,"<NCI>", part1);
          UpdateSlidePojo call2 = new UpdateSlidePojo(56,"<JSC>", part2);
          UpdateSlidePojo call3 = new UpdateSlidePojo(56,"<JSC%>", String.format("%.0f", persent) + "%");
      	  UpdateSlidePojo call5 = new UpdateSlidePojo(56,"<2YIFWSC>", IFwithout);
        
      	  UpdateSlidePojo call4 = new UpdateSlidePojo(11,"<2YEARIF>", IFwith);
          UpdateSlidePojo call6 = new UpdateSlidePojo(11,"<5YEARIF>", fiveyear); 
          UpdateSlidePojo call7 = new UpdateSlidePojo(56,"<2YEARIF>", IFwith);
          UpdateSlidePojo call8 = new UpdateSlidePojo(56,"<5YEARIF>", fiveyear); 
          
          	updateSlides.add(call1); 
        	updateSlides.add(call2); 
        	updateSlides.add(call3); 
        	updateSlides.add(call4); 
        	updateSlides.add(call5);
        	updateSlides.add(call6);
        	updateSlides.add(call7);
        	updateSlides.add(call8);
        	pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides);
} catch(Exception e1) {}
      	}
      }
      File myfile0 = new File(downloadFilepath + "\\JCR_category1.xlsx");
		if(myfile0.exists()){
			a = 5;
		}
	}
      driver.close();
      driver.quit();
} catch(Exception e1) {} 
        	}}	
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
