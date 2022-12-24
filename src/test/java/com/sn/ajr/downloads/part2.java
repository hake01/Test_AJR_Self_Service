package com.sn.ajr.downloads;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.slides.v1.Slides;
import com.sn.ajr.ProjGSheet;
import com.sn.ajr.ProjGSlide; 

public class part2 {
	public static WebDriver driver;
	@Test( timeOut = 180000 )
	public static void mainclass() throws InterruptedException, IOException, GeneralSecurityException {
		System.out.println("part2");	
		try {
		File currentDirFile = new File(".");
		String rootpath1 = currentDirFile.getAbsolutePath();
		String rootpath = rootpath1.replace(".", ""); 
		String livepath = rootpath + "batchfiles\\live.txt"; 
		String sheetidq = "";
    	File fa1 = new File(livepath);
    	if(fa1.exists()){
    		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
    	}else {
    		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
    	}
//		String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example"); 
    	String sheetname = "";
    	String JournalID = "";
    	String JournalName = "";
    	String pISSN = "";
    	String eISSN = "";
    	String Platform = ""; 
    	String JournalID1 = "";
    	String pptid = "";   
    	String uid = "";
		String uidp = "";
    	try {
    	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
		for (List<?> row : values) {
		sheetname = (String) row.get(0);
    	JournalID = (String) row.get(1);
    	JournalName = (String) row.get(2);
    	pISSN = (String) row.get(3);
    	eISSN = (String) row.get(4);
    	Platform = (String) row.get(5);
    	JournalID1 = (String) row.get(10);
    	pptid = (String) row.get(11);
		}
		List<List<Object>> values1 = sheetq.getCellData(sheetsService1, "Details!C4:E4",sheetidq); 
		for (List<?> row : values1) {
			uid = (String) row.get(1);
			uidp = (String) row.get(2);
		}
    	} catch (Exception e1) {}		
    	String sheetid = "";
    	if(!JournalID.contentEquals("")) {
    		
        	try {
        	List<List<Object>> values = sheetq.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
        	for (List<?> row : values) {
        	sheetid = (String) row.get(0);
        	}
        	 } catch (Exception e1) {}
        
        if(!sheetid.contentEquals("")) {  
    	ProjGSheet sheet = new ProjGSheet();
		Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
		ProjGSlide pSlide = new ProjGSlide();
		Slides service = pSlide.getSlideService("name");
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		driver = new ChromeDriver();
		String scopus = "";
		String QTS = "";
		String Retractionvalue = "";
		String h5index = "";
		String inddexvalue = "";
		String about = "";
		try {
        	List<List<Object>> values = sheet.getCellData(sheetsService1, sheetname + "!E2:J2",sheetid); 
        	for (List<?> row : values) {
        	scopus = (String) row.get(0);
    		QTS = (String) row.get(1);
    		Retractionvalue = (String) row.get(2);
    		h5index = (String) row.get(3);
    		inddexvalue = (String) row.get(4);
    		about = (String) row.get(5);
        	}
        	 } catch (Exception e1) {}
		
		// www.scopus.com	
		try {
		if(scopus.contentEquals("No")) {
			
					driver.get("https://www.scopus.com/sources.uri");
					driver.manage().window().maximize();
					new WebDriverWait(driver, Duration.ofSeconds(10)).until(
							ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='srcResultComboDrp-button']")));
					Thread.sleep(400);
					if (!driver.findElements(By.xpath("//*[@id=\"srcResultComboDrp-button\"]")).isEmpty()) {
						driver.findElement(By.xpath("//*[@id=\"srcResultComboDrp-button\"]/span[1]")).click();
						Thread.sleep(1000);
						driver.findElement(By.id("ui-id-4")).click();
						Thread.sleep(100);
						driver.findElement(By.xpath("//*[@class='inputTextLabel']")).click();
						Thread.sleep(100);
						driver.findElement(By.xpath("//*[@id='sourceResultSearchInp']//input")).sendKeys(eISSN);
						Thread.sleep(500);
						driver.findElement(By.id("searchTermsSubmit")).click();
						Thread.sleep(5000);
						if (driver.findElements(By.xpath("(//*[@title='View details for this source.'])[1]")).isEmpty()) {
							driver.findElement(By.xpath("//*[@class='inputTextLabel']")).click();
							Thread.sleep(100);
							driver.findElement(By.xpath("//*[@id='sourceResultSearchInp']//input")).sendKeys(pISSN);
							Thread.sleep(500);
							driver.findElement(By.id("searchTermsSubmit")).click();
							Thread.sleep(5000);
						}
						if (!driver.findElements(By.xpath("(//*[@title='View details for this source.'])[1]")).isEmpty()) {
							WebElement viewtext = driver.findElement(By.xpath("(//*[@title='View details for this source.'])[1]"));
							JavascriptExecutor jsq = (JavascriptExecutor) driver;
							jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);							
							driver.findElement(By.xpath("(//*[@title='View details for this source.'])[1]")).click();
							Thread.sleep(1000);
							WebElement viewtext1 = driver.findElement(By.xpath("//*[@id='CSCategoryTBody']"));
							JavascriptExecutor jsq1 = (JavascriptExecutor) driver;
							jsq1.executeScript("arguments[0].scrollIntoView(true);", viewtext1);	
							new WebDriverWait(driver, Duration.ofSeconds(10)).until(
							ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='CSCategoryTBody']")));
							Thread.sleep(500);
							String CSCAtegory = driver.findElement(By.xpath("//*[@id='CSCategoryTBody']")).getText();
							String csCalculation = driver.findElement(By.xpath("//*[@id='csCalculation']//div[2]//div[2]"))
									.getText();
							sheet.updateCellvalue(sheetsService, sheetname + "!F215", csCalculation + " calcate " + '\n' + CSCAtegory,sheetid);
						}
					}
		}
				} catch (Exception e1) {}
		
// QTS Disapproval rate
		try {
		if(QTS.contentEquals("No")) {
		driver.get("https://qts.springernature.com/journal_production_view.php?jnum=" + JournalID + "&year=2021");
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='username']")));
			Thread.sleep(100);
			driver.findElement(By.name("username")).sendKeys(uid);
			driver.findElement(By.name("password")).sendKeys(uidp);
			driver.findElement(By.name("btnLogin")).click();
			Thread.sleep(100);
			WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='chart5']")));
			Thread.sleep(100);
			if (!driver.findElements(By.xpath("//*[@id='chart5']")).isEmpty()) {
				String dvalue = driver.findElement(By.xpath("(//*[@id='chart5']//*[@text-anchor='middle'])[2]"))
						.getText();
				String dvalue1 = dvalue.replaceAll("\n", " ");
				Thread.sleep(500);
			sheet.updateCellvalue(sheetsService, sheetname + "!F6", dvalue1, sheetid);
			}
		}
		} catch (Exception e1) {}
		
// Retractionvalue
		try {
		if(Retractionvalue.contentEquals("No")) {
		
			driver.get("http://retractiondatabase.org/RetractionSearch.aspx?");
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='txtSrchJournal']")));
			driver.findElement(By.name("txtSrchJournal")).sendKeys(JournalName);
			Thread.sleep(1000);
			driver.findElement(By.name("txtFromDate")).sendKeys("01/01/2021");
			driver.findElement(By.name("txtToDate")).sendKeys("12-31-2021");
			driver.findElement(By.name("drpNature")).sendKeys("Retraction");
			driver.findElement(By.name("btnSearch")).click();
			Thread.sleep(1200);
			if (!driver.findElements(By.xpath("//*[@class='totalItems']")).isEmpty()) {
				String evalue = driver.findElement(By.xpath("//*[@class='totalItems']")).getText();
				sheet.updateCellvalue(sheetsService, sheetname + "!F4", evalue,sheetid);
			}}
		} catch (Exception e1) {}
		
// h5index
		try {
		if(h5index.contentEquals("No")) {
			driver.get("https://scholar.google.com/citations?view_op=top_venues&hl=en&vq=en");
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='gs_hdr_sre']")));
			driver.findElement(By.xpath("//*[@id='gs_hdr_sre']")).click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//*[@placeholder='Search publications']")).sendKeys(JournalName);
			driver.findElement(By.xpath("//*[@type='submit']")).click();
			Thread.sleep(1000);
			String JournalName1 = JournalName;
			if (driver.findElements(By.xpath("//*[@id='gsc_mvt_table']//tr[2]//td[2]")).isEmpty()) {
				if (JournalName.startsWith("The ")) {
					JournalName1 = JournalName.replace("The ", "");
				}
				if (JournalName.startsWith("Der ")) {
					JournalName1 = JournalName.replace("Der ", "");
				}
				new WebDriverWait(driver, Duration.ofSeconds(10)).until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath("(//*[@aria-label='Search'])[1]")));
				driver.findElement(By.xpath("(//*[@aria-label='Search'])[1]")).click();
				driver.findElement(By.xpath("(//*[@aria-label='Search'])[1]")).clear();
				Thread.sleep(1000);
				driver.findElement(By.xpath("//*[@placeholder='Search publications']")).sendKeys(JournalName1);
				driver.findElement(By.xpath("//*[@type='submit']")).click();
				Thread.sleep(1000);
			}

			for (int i12 = 2; i12 < 25; i12++) {
				String indexcheck = driver.findElement(By.xpath("//*[@id='gsc_mvt_table']//tr[" + i12 + "]//td[2]"))
						.getText();
				if (indexcheck.contentEquals(JournalName1)) {
					driver.findElement(By.xpath("//*[@id='gsc_mvt_table']//tr[" + i12 + "]//td[3]")).click();
					i12 = 25;
				}
			}

			String h5_index = driver.findElement(By.xpath("//*[@id='gsc_mp_desc']")).getText();
			sheet.updateCellvalue(sheetsService, sheetname + "!A44", h5_index,sheetid);
		}} catch (Exception e1) {}
		
// index
		try {
			if(inddexvalue.contentEquals("No")) {
			driver.get("https://tools.springernature.com");
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='username']")));
			driver.findElement(By.xpath("//*[@name='username']")).sendKeys(uid);
			driver.findElement(By.xpath("//*[@name='password']")).sendKeys(uidp);
			driver.findElement(By.xpath("//*[@type='submit']")).click();
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='ani_reports']")));
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@id='ani_reports']")).click();
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='s2id_product_title']")));
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@id='s2id_product_title']")).click();
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@id='s2id_product_title']//input[1]")).sendKeys(JournalName);
			new WebDriverWait(driver, Duration.ofSeconds(10))
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='select2-drop']//li[1]")));
			Thread.sleep(700);
			String valuecheck = "";
			for (int i1 = 1; i1 < 10; i1++) {
				valuecheck = driver.findElement(By.xpath("//*[@id='select2-drop']//li[" + i1 + "]")).getText();
				if (valuecheck.contentEquals(JournalName)) {
					driver.findElement(By.xpath("//*[@id='select2-drop']//li[" + i1 + "]")).click();
					i1 = 10;
				}
			}
			driver.findElement(By.linkText("Fetch Report")).click();
			Thread.sleep(2000);
			String inddex = driver.findElement(By.xpath("//*[@class=\"odd\"]//td[5]")).getText();
			System.out.println(inddex);
			
			sheet.updateCellvalue(sheetsService, sheetname + "!B37", " inddex" + " = " + inddex,sheetid);
			}} catch (Exception e1) {}
		
	// about this journal and editorial board
		try {
		if(about.contentEquals("No")) {
			String aboutj = "";
			String editornames = "";
			String coverlink = "";
			try {
				driver.get("https://www.springer.com/journal/" + JournalID);
				driver.manage().window().maximize();
				Thread.sleep(5000);
				coverlink = driver.getCurrentUrl();
				if (coverlink.contentEquals("https://www.springer.com/journal/" + JournalID)) {
					coverlink = "www.springer.com/" + JournalID;
				}
				if (!driver.findElements(By.xpath("//*[@class='cc-banner__button-break']")).isEmpty()) {
					driver.findElement(By.xpath("//*[@class='cc-banner__button-break']")).click();
				}
			} catch (Exception e1) {}

			try {
//			Springer Open				
				if (!driver.findElements(By.xpath("//*[@id='siteNavigation']//li[1]")).isEmpty()) {
					driver.findElement(By.xpath("//*[@id='siteNavigation']//li[1]")).click();
					aboutj = driver.findElement(By.xpath("(//*[@class='c-page-layout__main']//div[1])[1]")).getText();
					driver.findElement(By.xpath("//*[@id='siteNavigationChildren']//li[2]/a")).click();
					editornames = driver.findElement(By.xpath("//*[@id=\"editorialboard\"]")).getText();
				}
			} catch (Exception e1) {
				System.out.println("IOException");
			}
			try {
//			Palgrave Macmillan
				if (!driver.findElements(By.xpath("//*[@class='journal-navigation']//li[3]")).isEmpty()) {
					aboutj = driver.findElement(By.xpath("(//*[@class='live-area']//div[3])[1]")).getText();
					driver.findElement(By.xpath("//*[@class='journal-navigation']//li[3]")).click();
					editornames = driver.findElement(By.xpath("//*[@class='article-body']")).getText();
				}
			} catch (Exception e1) {
				System.out.println("IOException");
			}
			try {
//			springer
				if (!driver.findElements(By.xpath("//*[@class='app-promo-text app-promo-text--keyline']")).isEmpty()) {
					if (!driver.findElements(By.xpath("//*[@class=\"app-promo-text__button\"]")).isEmpty()) {
						driver.findElement(By.xpath("//*[@class=\"app-promo-text__button\"]")).click();
					}
				}
				if (!driver.findElements(By.xpath("//*[@class='app-promo-text app-promo-text--keyline']")).isEmpty()) {
					aboutj = driver.findElement(By.xpath("//*[@class='app-promo-text app-promo-text--keyline']"))
							.getText();
				} else {
					driver.findElement(By.xpath("//*[@data-test='aimsAndScopeLink']")).click();
					aboutj = driver.findElement(By.xpath("//*[@id='aimsAndScope']")).getText();
					driver.findElement(By.xpath("//*[@data-test='journal-homepage-breadcrumb']")).click();
				}
				driver.findElement(By.xpath("//*[@data-test='editorialboardLink']")).click();
				editornames = driver.findElement(By.xpath("//*[@id='editorialboard']")).getText();
			} catch (Exception e1) {
				System.out.println("IOException");
			}
			try {
				if (!driver.findElements(By.xpath("//*[@data-track-category=\"About\"]")).isEmpty()) {
					driver.findElement(By.xpath("//*[@data-track-category=\"About\"]")).click();
					aboutj = driver.findElement(By.xpath("(//*[@class='c-page-layout__main']//div[1])[1]")).getText();
					driver.findElement(By.xpath("//*[@id=\"siteNavigationChildren\"]//li[2]")).click();
					editornames = driver.findElement(By.id("editorialboard")).getText();
				}
				if (!driver.findElements(By.xpath("//*[@data-test='menu-button--about-the-journal']")).isEmpty()) {
					driver.findElement(By.xpath("//*[@data-test='menu-button--about-the-journal']")).click();
					driver.findElement(By.xpath("//*[@data-track-action='journal information']")).click();
					aboutj = driver.findElement(By.xpath("//*[@data-container-type='body-content']")).getText();
					driver.findElement(By.xpath("//*[@data-test='menu-button--about-the-journal']")).click();
					driver.findElement(By.xpath("//*[@data-track-action='about the editors']")).click();
					if (!driver.findElements(By.xpath("//*[@data-track-action='editorial board']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@data-track-action='editorial board']")).click();
					}
					editornames = driver.findElement(By.xpath("(//*[@id='content']//div[2])[1]")).getText();
				}
			} catch (Exception e1) {
				System.out.println("IOException");
			}
			
			try {
//			BMC
				if (Platform.contentEquals("BMC")) {
					if (!driver.findElements(By.xpath("//*[@data-track-category='About']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@data-track-category='About']")).click();
						aboutj = driver.findElement(By.xpath("(//*[@class='c-page-layout__main']//div[1])[1]"))
								.getText();
						driver.findElement(By.xpath("//*[@id='siteNavigationChildren']//li[2]")).click();
						editornames = driver.findElement(By.id("editorialboard")).getText();
					}
					if (!driver.findElements(By.xpath("//*[@data-test='menu-button--about-the-journal']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@data-test='menu-button--about-the-journal']")).click();
						driver.findElement(By.xpath("//*[@data-track-action='journal information']")).click();
						aboutj = driver.findElement(By.xpath("//*[@data-container-type='body-content']")).getText();
						driver.findElement(By.xpath("//*[@data-track-action='about the editors']")).click();
						if (!driver.findElements(By.xpath("//*[@data-track-action='editorial board']")).isEmpty()) {
							driver.findElement(By.xpath("//*[@data-track-action='editorial board']")).click();
						}
						editornames = driver.findElement(By.xpath("(//*[@id='content']//div[2])[1]")).getText();
					}
				}
			} catch (Exception e1) {
				System.out.println("IOException");
			}

			try {
//				Nature
				if (Platform.contentEquals("Nature")) {
					if (!driver.findElements(By.xpath("//*[@data-test='menu-button--about-the-journal']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@data-test='menu-button--about-the-journal']")).click();
						driver.findElement(By.xpath("//*[@data-track-action='journal information']")).click();
						Thread.sleep(100);
//					aboutj = driver.findElement(By.xpath("(//*[@data-container-type='body-content']//p)[1]")).getText();
						aboutj = driver.findElement(By.xpath("//*[@data-container-type='body-content']")).getText();

						driver.findElement(By.xpath("//*[@data-test='menu-button--about-the-journal']")).click();
						driver.findElement(By.xpath("//*[@data-track-action='about the editors']")).click();
						Thread.sleep(100);

						if (!driver.findElements(By.xpath("//*[@data-container-type='section-navigation']//li[2]"))
								.isEmpty()) {
							driver.findElement(By.xpath("//*[@data-container-type='section-navigation']//li[2]"))
									.click();
						}
						editornames = driver.findElement(By.xpath("(//*[@id='content']//div[2])[1]")).getText();
					}
				}
			} catch (Exception e1) {
				System.out.println("IOException");
			}
			try {
				String[] parts = aboutj.split("Impact Factor");
				aboutj = parts[0];

				sheet.updateCellvalue(sheetsService, sheetname + "!W1", aboutj,sheetid);
				sheet.updateCellvalue(sheetsService, sheetname + "!X1", editornames,sheetid);
				pSlide.updatePresentationData(pptid, service, "www.springer.com/JournalID", coverlink);
				pSlide.updatePresentationData(pptid, service, "JournalID", JournalID1);					
				Thread.sleep(1000);
			} catch (Exception e1) {System.out.println("IOException");}
		}} catch (Exception e1) {}
		
		
    	}
        }
		} catch (Exception e1) {}
		try {
			driver.close();
			driver.quit();
	} catch (Exception e1) {}
	}
	  @AfterTest
	  public void afterTest() {	 
		  try {
				driver.close();
				driver.quit();
			} catch(Exception e1) {System.out.println("Exceptionbex");}
	  }
}