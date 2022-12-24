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
public class part1 {
	public static WebDriver driver;
	@Test( timeOut = 180000 )
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
    	String sheetname = "";
    	String JournalID = "";
    	String pISSN = "";
    	String eISSN = "";
    	String Covername = "";
    	String shorttitle = "";
    	try {
    	List<List<Object>> values = sheet1.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
		for (List<?> row : values) { 
    	sheetname = (String) row.get(0);
    	JournalID = (String) row.get(1);
    	pISSN = (String) row.get(3);
    	eISSN = (String) row.get(4);
    	shorttitle = (String) row.get(6);  
    	Covername = (String) row.get(8);
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
		ProjGSheet sheet = new ProjGSheet();
		Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");    	

		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--disable-gpu");
		driver = new ChromeDriver(options);

		String classifications = "";
		String PubMed = ""; 
		try {
        	List<List<Object>> values = sheet.getCellData(sheetsService1, sheetname + "!K2:N2",sheetid); 
        	for (List<?> row : values) {
        		classifications = (String) row.get(2);
        		PubMed = (String) row.get(3);
        	}
        	 } catch (Exception e1) {}
	
//		Cover
		try {
		for (int a = 1; a < 5; a++) {
		File fa01 = new File(downloadFilepath + "Cover.png");
		if(!fa01.exists()){		
		driver.get("https://covers.springernature.com/search/CoverSearch.html");
		driver.manage().window().maximize();	
			Thread.sleep(2000);
			driver.findElement(By.xpath("//body/table//tr[3]//td[1]//textarea")).sendKeys(Covername);
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@value=\"journals_single_issue/\"]")).click();
			Thread.sleep(100);
			driver.findElement(By.xpath("//*[@id='show']")).click();
			Thread.sleep(7000);
			if (driver.findElement(By.xpath("//*[@id='message']")).getText().contentEquals("Success: 1")) {
				WebElement viewtext = driver.findElement(By.xpath("//*[@id='previewimg_0:0']"));
				JavascriptExecutor jsq = (JavascriptExecutor) driver;
				jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);				
				Thread.sleep(500);
				driver.findElement(By.xpath("//*[@id='previewimg_0:0']")).click();
				Thread.sleep(20000);
				final BufferedImage tif = ImageIO.read(new File(downloadFilepath + Covername + ".tif"));
				ImageIO.write(tif, "png", new File(downloadFilepath + "Cover.png"));
			} else {
				driver.findElement(By.xpath("//body/table//tr[3]//td[1]//textarea")).clear();
				driver.findElement(By.xpath("//body/table//tr[3]//td[1]//textarea")).sendKeys(JournalID);
				driver.findElement(By.xpath("//*[@value='journals/']")).click();
				Thread.sleep(100);
				driver.findElement(By.xpath("//*[@id='show']")).click();
				Thread.sleep(7000);
				if (driver.findElement(By.xpath("//*[@id='message']")).getText().contentEquals("Success: 1")) {
					WebElement viewtext = driver.findElement(By.xpath("//*[@id='previewimg_0:0']"));
					JavascriptExecutor jsq = (JavascriptExecutor) driver;
					jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);				
					Thread.sleep(500);
					driver.findElement(By.xpath("//*[@id='previewimg_0:0']")).click();
					Thread.sleep(15000);
					final BufferedImage tif = ImageIO.read(new File(downloadFilepath + JournalID + ".tif"));
					ImageIO.write(tif, "png", new File(downloadFilepath + "Cover.png"));
				}
			}
		}
		}
		} catch (Exception e1) {}
		
		// classifications0029-599X 256
		try {
		if(classifications.contentEquals("No")) {  
			String issnlink = "https://SPRINGLON:420124@mathscinet.ams.org/mathscinet/search/journals.html?journalName="
							+ eISSN + "&Submit=Search";
					driver.get(issnlink);
					Thread.sleep(5000);
						String vaar0 = driver.findElement(By.xpath("//*[@id='content']/p"))
								.getText();
						if (!vaar0.contentEquals("No matches found")) {
					try {
						for (int vaar = 1; vaar < 5; vaar++) {
							try {
								String vaar1 = driver.findElement(By.xpath("(//*[@class='jheadline'])[" + vaar + "]//span"))
										.getText();
								if (vaar1.contentEquals("[Indexed cover-to-cover; Reference List Journal]")) {
									driver.findElement(By.xpath("(//*[@class='jheadline'])[" + vaar + "]//a")).click();
								}
								if (vaar1.isEmpty()) {
									driver.findElement(By.xpath("(//*[@class='jheadline'])[" + vaar + "]//a")).click();
								}
							} catch (Exception e1) {
								driver.findElement(By.xpath("(//*[@class='jheadline'])[" + vaar + "]//a")).click();
							}

						}
					} catch (Exception e1) {}
					try {
						if (!driver.findElements(By.xpath("(//*[@class=\"nav-item\"])[2]")).isEmpty()) {
							driver.findElement(By.xpath("(//*[@class=\"nav-item\"])[2]")).click();
							Thread.sleep(1000);
							String classificationmcq = driver
									.findElement(By.xpath("(//*[@class='table table-striped table-sm'])[3]")).getText();
							if (!driver.findElements(By.xpath("//*[@class=\"hide-xxs\"]")).isEmpty()) {
								driver.findElement(By.xpath("//*[@class=\"hide-xxs\"]")).click();
							}
							Thread.sleep(1000);
							if (!driver.findElements(By.xpath("//*[@class='tabs citations']//li[4]")).isEmpty()) {
								driver.findElement(By.xpath("(//*[@class='tabs citations']//li[3])[1]")).click();

								driver.findElement(By.xpath("//*[@id='year-select']")).click();
								driver.findElement(By.xpath("//*[@value='2021']")).click();
								driver.findElement(By.xpath("//*[@class='form-group form-row indent-xs']//button")).click();
								Thread.sleep(7000);
								driver.findElement(By.xpath("//*[@data-title='Download plot as a png']")).click();
							}
							Thread.sleep(1000);
							driver.findElement(By.xpath("(//*[@role='tablist'])[4]//li[2]")).click();
							Thread.sleep(1000);
							String classificationallmcq = driver
									.findElement(By.xpath("(//*[@class='table table-striped table-sm'])[4]")).getText();
							String classificationtab = driver
									.findElement(By.xpath("(//*[@class='table table-striped table-sm'])[9]")).getText();

								sheet.updateCellvalue(sheetsService, sheetname + "!F207", "aclasstab" + classificationtab,sheetid);
								sheet.updateCellvalue(sheetsService, sheetname + "!G215", " cmcq" + classificationmcq,sheetid);
								sheet.updateCellvalue(sheetsService, sheetname + "!H215", " allmcq" + classificationallmcq,sheetid);
							Thread.sleep(1000);
						}
					} catch (Exception e1) {}					
					}
		}		} catch (Exception e1) {}
		
		 //PubMed
		try {
		if(PubMed.contentEquals("No")) {
			if (!shorttitle.isEmpty()) {
				driver.get("https://account.ncbi.nlm.nih.gov/?back_url=https%3A//pubmed.ncbi.nlm.nih.gov/");
				driver.manage().window().maximize();
				Thread.sleep(5000);
				if (!driver.findElements(By.linkText("NCBI Account")).isEmpty()) {
					driver.findElement(By.linkText("NCBI Account")).click();
					Thread.sleep(2000);
					if (!driver.findElements(By.xpath("//*[@name='username']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@name='username']")).sendKeys("Springer_pmc");
						Thread.sleep(100);
						driver.findElement(By.xpath("//*[@name='password']")).sendKeys("openchoice");
						Thread.sleep(500);
						driver.findElement(By.className("submit-button")).click();
					}
					Thread.sleep(1000);
					driver.get(
							"https://www.ncbi.nlm.nih.gov/pmc/publisherservices/statistics2/reports/2?s=1&sc=2&sct=3&pb=&j=960&pb_pmc=springer&j_pmc=pharmsci&e=&a=");
					Thread.sleep(12000);
				}
				if (!driver.findElements(By.xpath("//*[@class=\"ant-form ant-form-inline\"]//div[3]//input"))
						.isEmpty()) {
				} else {
					driver.get(
							"https://www.ncbi.nlm.nih.gov/pmc/publisherservices/statistics2/reports/2?s=1&sc=2&sct=3&pb=&j=960&pb_pmc=springer&j_pmc=pharmsci&e=&a=");
					Thread.sleep(12000);
				}
				if (!driver.findElements(By.xpath("//*[@class='ant-form ant-form-inline']//div[3]//input")).isEmpty()) {
					driver.findElement(By.xpath("//*[@arial-label='Close this window']//span")).click();
				}
				if (!driver.findElements(By.xpath("//*[@class='ant-form ant-form-inline']//div[3]//input")).isEmpty()) {
					driver.findElement(By.xpath("//*[@class='ant-form ant-form-inline']//div[3]//input"))
							.sendKeys(shorttitle);
					if (!driver.findElements(By.xpath("//*[@title='" + shorttitle + "']")).isEmpty()) {
						driver.findElement(By.xpath("//*[@title='" + shorttitle + "']")).click();

						Thread.sleep(5000);
						driver.findElement(By.xpath("//*[@class='ant-btn ant-dropdown-trigger ant-btn-primary']"))
								.click();
						Thread.sleep(500);
						driver.findElement(By.xpath(
								"//*[@class='ant-dropdown-menu ant-dropdown-menu-light ant-dropdown-menu-root ant-dropdown-menu-vertical']//li[2]"))
								.click();
						Thread.sleep(15000);
						driver.close();
						// csv to xlsx
						String csvFileAddress111 = downloadFilepath + "PMCStatistics.csv";  
						String xlsxFileAddress111 = downloadFilepath + "PMCStatistics.xlsx"; 
						try {
							@SuppressWarnings("resource")
							XSSFWorkbook workBook = new XSSFWorkbook();
							XSSFSheet sheet111 = workBook.createSheet("sheet1");
							String currentLine = null;
							int RowNum = -1;
							BufferedReader br = new BufferedReader(new FileReader(csvFileAddress111));
							while ((currentLine = br.readLine()) != null) {
								String str[] = currentLine.split(",");
								RowNum++;
								XSSFRow currentRow = sheet111.createRow(RowNum);
								for (int i = 0; i < str.length; i++) {
									currentRow.createCell(i).setCellValue(str[i]);
								}
							}
							FileOutputStream fileOutputStream = new FileOutputStream(xlsxFileAddress111);
							workBook.write(fileOutputStream);
							fileOutputStream.close();
							br.close();
							System.out.println("Done");
						} catch (Exception ex) {
							System.out.println(ex.getMessage() + "Exception in try");
						}
						FileInputStream inputStream = null;
						try {
							inputStream = new FileInputStream(new File(downloadFilepath + "PMCStatistics.xlsx"));
						} catch (FileNotFoundException e1) { 
							e1.printStackTrace();
						}
						XSSFWorkbook workbook1 = null;
						try {
							workbook1 = new XSSFWorkbook(inputStream);
						} catch (IOException e1) { 
							e1.printStackTrace();
						}
						XSSFSheet sheet111 = workbook1.getSheetAt(0);
						String row0 = "";
						for (int j = 1; j <= 24; j++) {
							if (!(sheet111.getRow(j) == null)) {
								XSSFRow row = sheet111.getRow(j);
								XSSFCell cell = row.getCell(0);
								if (cell != null) { 
									row0 = row0 + '\n' + cell;
								}
							}
						}
						String row1 = "";
						for (int j = 1; j <= 24; j++) {
							if (!(sheet111.getRow(j) == null)) {
								XSSFRow row = sheet111.getRow(j);
								XSSFCell cell = row.getCell(4);
								if (cell != null) {
									row1 = row1 + '\n' + cell;
								}
							}
						}

						String row2 = "";
						for (int j = 1; j <= 24; j++) {
							if (!(sheet111.getRow(j) == null)) {
								XSSFRow row = sheet111.getRow(j);
								XSSFCell cell = row.getCell(5);
								if (cell != null) {
									row2 = row2 + '\n' + cell;
								}
							}
						}

						String row3 = "";
						for (int j = 1; j <= 24; j++) {
							if (!(sheet111.getRow(j) == null)) {
								XSSFRow row = sheet111.getRow(j);
								XSSFCell cell = row.getCell(7);
								if (cell != null) {
									row3 = row3 + '\n' + cell;
								}
							}
						}
						sheet.updateCellvalue(sheetsService, sheetname + "!R188",
								"APMCStatisticsA|" + row0 + '\n' + "|" + row1 + '\n' + "|" + row2 + '\n' + "|" + row3,
								sheetid);
					} else {}
				}
			}
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
