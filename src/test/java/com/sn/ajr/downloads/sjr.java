package com.sn.ajr.downloads;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.HashMap;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets; 
import com.sn.ajr.ProjGSheet;
public class sjr {
	public static WebDriver driver;
	@Test( timeOut = 70000 )
	public static void classname() throws InterruptedException, IOException, GeneralSecurityException {
		System.out.println("part1");
		try {
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
		ProjGSheet sheet1 = new ProjGSheet();
		Sheets sheetsService1 = sheet1.getSheetsService("Google Sheets Example"); 
    	String JournalID = "";
    	String pISSN = "";
    	String eISSN = "";
    	try {
    	List<List<Object>> values = sheet1.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
		for (List<?> row : values) { 
    	JournalID = (String) row.get(1);
    	pISSN = (String) row.get(3);
    	eISSN = (String) row.get(4);
    	}
    	} catch (Exception e1) {}
		
    	if(!JournalID.contentEquals("")) {  
    	String sheetid = "";
    	try {
    	List<List<Object>> values = sheet1.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
    	for (List<?> row : values) {
    	sheetid = (String) row.get(0);
    	}
    	 } catch (Exception e1) {}
    	if(!sheetid.contentEquals("")) {  
		 
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--disable-gpu");
		driver = new ChromeDriver(options);

		// SJR
				try {
				for (int a1 = 1; a1 < 5; a1++) {
				File fa02 = new File(downloadFilepath + "SJR.png");
				if(!fa02.exists()){
					driver.get("https://www.scimagojr.com/journalsearch.php");
					driver.manage().window().maximize();
					Thread.sleep(3000);
					if (!driver.findElements(By.xpath("//*[@class='ccb__button']//button[2]")).isEmpty()) {
						driver.findElement(By.xpath("//*[@class='ccb__button']//button[2]")).click();
					}
					System.out.println("checksjr0");
					driver.findElement(By.xpath("(//*[@id='searchinput'])[2]")).sendKeys(eISSN);
					Thread.sleep(3000);
					System.out.println("checksjr001"); 
					if (!driver.findElements(By.xpath("//*[@id='searchbutton']")).isEmpty()) {
						driver.findElement(By.xpath("(//*[@id='searchbutton'])[2]")).click();
					}
					System.out.println("checksjr1");
					Thread.sleep(5000);
					if (driver.findElements(By.xpath("//*[@class='jrnlname']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@id='searchinput']")).sendKeys(pISSN);
						driver.findElement(By.xpath("//*[@id='searchbutton']")).click();
						Thread.sleep(2000);
					}
					System.out.println("checksjr2");
//					if (!driver.findElements(By.xpath("//*[@class='search_results']//a")).isEmpty()) {
//						driver.findElement(By.xpath("//*[@class='search_results']//a")).click();
					if (!driver.findElements(By.xpath("//*[@class='jrnlname']")).isEmpty()) {
						System.out.println("checksjr3");
						driver.findElement(By.xpath("//*[@class='jrnlname']")).click();
						System.out.println("checksjr4");
					}
					Thread.sleep(4000);
					if (!driver.findElements(By.xpath("//*[@class='cc_btn cc_btn_accept_all']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@class='cc_btn cc_btn_accept_all']")).click();
					}
					
					if (!driver.findElements(By.xpath("//*[@id='svgquartiles']")).isEmpty()) {
						System.out.println("checksjr5");
						WebElement SJRlogo = driver.findElement(By.xpath("//*[@id='svgquartiles']"));
						JavascriptExecutor jsq = (JavascriptExecutor) driver;
						jsq.executeScript("arguments[0].scrollIntoView(true);", SJRlogo);
						Thread.sleep(1000);
						File f = SJRlogo.getScreenshotAs(OutputType.FILE);
						FileUtils.copyFile(f, new File(downloadFilepath + "SJR.png"));
						Thread.sleep(4000);
					}
				}
				}} catch (Exception e1) {}
				  try {
						driver.close();
						driver.quit();
					} catch(Exception e1) {}
    	}}
		} catch (Exception e1) {}
		}
	
  @BeforeTest
  public void beforeTest() {
  }

  @AfterTest
  public void afterTest() {
	  try {
			driver.close();
			driver.quit();
		} catch(Exception e1) {}
  }

}
