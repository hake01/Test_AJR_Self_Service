package com.sn.ajr.altmetric;
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.slides.v1.Slides;
import com.sn.ajr.ProjGSheet;
import com.sn.ajr.ProjGSlide;
import com.sn.ajr.UpdateSlidePojo;
import com.spire.xls.OrderBy;
import com.spire.xls.SortComparsionType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.sorting.SortColumn;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List; 
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;

public class AltmetricMaster {
	public static WebDriver driver;
	public static WebDriver driver1;
	@Test( timeOut = 240000 )
  public static void classname() throws InterruptedException, IOException, GeneralSecurityException { 
	System.out.println("alt start");
	  final String APPLICATION_NAME = "Google Slides API Java Quickstart"; 
	File currentDirFile = new File(".");
	String rootpath1 = currentDirFile.getAbsolutePath();
	String rootpath = rootpath1.replace(".", "");
	String downloadFilepath = rootpath + "DownloadingTempfiles\\";
	String livepath = rootpath + "batchfiles\\live.txt"; 
	String sheetidq = "";
	File f0 = new File(livepath);
	if(f0.exists()){
		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
	}else {
		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
	}
//	String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
	ProjGSheet sheetq = new ProjGSheet();
	Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example");   
  	String pISSN = "";
  	String eISSN = "";
	String altid = "";
	String altp = "";
	String sheetname = "";
	String pptid = ""; 
  	try {
  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
		for (List<?> row : values) {
	sheetname = (String) row.get(0);
	pISSN = (String) row.get(3);
  	eISSN = (String) row.get(4);
  	pptid = (String) row.get(11); 
		}
		List<List<Object>> values1 = sheetq.getCellData(sheetsService1, "Details!C6:E6",sheetidq); 
		for (List<?> row : values1) {
			altid = (String) row.get(0);
			altp = (String) row.get(2);
		}
  	} catch (Exception e1) {}
	
	if (eISSN.contentEquals("")){
		eISSN = pISSN;
	}
	if (eISSN.contentEquals("-")){
		eISSN = pISSN;
	}
	
  	if(!eISSN.contentEquals("")) {  		
  		String sheetid = "";
    	try {
    	List<List<Object>> values = sheetq.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
    	for (List<?> row : values) {
    	sheetid = (String) row.get(0);
    	}
    	 } catch (Exception e1) {}
    	if(!sheetid.contentEquals("")) {
  			String checkalt = "";
        	try {
        	List<List<Object>> values = sheetq.getCellData(sheetsService1, sheetname + "!M35:M35",sheetid); 
        	for (List<?> row : values) {
        		checkalt = (String) row.get(0);
        	}
        	 } catch (Exception e1) {}
        if(checkalt.contentEquals("")) {
		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--disable-gpu");
		driver = new ChromeDriver(options);
		String pattern = "yyyy-MM-dd";
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
		String date = simpleDateFormat.format(new Date());

	 	ProjGSlide pSlide = new ProjGSlide();
	  	Slides service = pSlide.getSlideService(APPLICATION_NAME);
	  	List<UpdateSlidePojo> updateSlides = new ArrayList<>();
		
	  	ProjGSheet sheet0 = new ProjGSheet();
		Sheets sheetsService = sheet0.getSheetsService("Google Sheets Example");
	  	
		String xlsxFileAddress = "Altmetric - Mentions_2019.xlsx";
		String xlsxFileAddress1 = "Altmetric - Mentions_2020.xlsx";
		String xlsxFileAddress11 = "Altmetric - Mentions_2021.xlsx";
		
		int firstrows = 1;
		int secondrows = 1; 
		String resa = "";
		String resb = "";
		String resc = "";
		String resd = "";
		String checkvalues = "";
		
		   try {
	        	File directoryq = new File(downloadFilepath);
	            File[] filesq = directoryq.listFiles();
	            for (File fq : filesq)           	
	            {
	            	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
	            		String oldfilename = fq.getName(); 
	                	File myfile = new File(downloadFilepath + oldfilename);
	                	myfile.delete();                	
	                }
	            }
	        } catch (Exception ex) {}
		
//Altmetric
//2019
for (int a = 1; a < 5; a++) {
File fa01 = new File(downloadFilepath + xlsxFileAddress);
if(!fa01.exists()){
try {
driver.get("https://www.altmetric.com/explorer/mentions?mentioned_after=2019-01-01&mentioned_before=2019-12-31");
driver.manage().window().maximize();
Thread.sleep(3000);
if(!driver.findElements(By.xpath("//*[@type='email']")).isEmpty()){
	driver.findElement(By.xpath("//*[@type='email']")).sendKeys(altid);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(2500);
	driver.findElement(By.xpath("//*[@type='password']")).sendKeys(altp);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(5000);
}
String checkoutput = "";
String checkoutputresult = "";
if(!driver.findElements(By.className("edit-button")).isEmpty()){
	driver.findElement(By.className("edit-button")).click();
	Thread.sleep(1000);
	WebElement viewtext = driver.findElement(By.xpath("//*[@id='journal_id']"));
	JavascriptExecutor jsq = (JavascriptExecutor) driver;
	jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);
	driver.findElement(By.id("journal_id")).sendKeys(eISSN);
	Thread.sleep(4000);
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='AutocompleteFieldResults']")));
	Thread.sleep(100);
	checkoutput = driver.findElement(By.className("AutocompleteFieldResults")).getText();	
	if(checkoutput.indexOf("o results found") > 0){
		checkoutput = "";
		driver.findElement(By.id("journal_id")).clear();
		Thread.sleep(2500);
		driver.findElement(By.id("journal_id")).sendKeys(pISSN);
		Thread.sleep(5000);
		checkoutput = driver.findElement(By.className("AutocompleteFieldResults")).getText();
	}
	if(checkoutput.indexOf("o results found") < 0){
	driver.findElement(By.className("AutocompleteFieldResults")).click();
	Thread.sleep(1000);
	if(!driver.findElements(By.className("AutocompleteField-item-list-item-remove-button")).isEmpty()){
		driver.findElement(By.className("execute")).click();
		new WebDriverWait(driver, Duration.ofSeconds(500))
		        .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='summary']//span")));
		Thread.sleep(100);
	 	if(!driver.findElements(By.xpath("//*[@class='summary']//span")).isEmpty()){
			checkoutputresult = driver.findElement(By.xpath("//*[@class='summary']//span")).getText();}
	 	Thread.sleep(2500);
	 	if(checkoutputresult.indexOf("no mentions") < 0){
		driver.findElement(By.xpath("//*[@class=\"export-option-selector\"]")).click();
		driver.findElement(By.partialLinkText("Download results as CSV (UTF-8)")).click();
		//driver.findElement(By.partialLinkText("Download results as CSV")).click();
		Thread.sleep(10000);
		
//csv to xls
String csvFileAddress = downloadFilepath +  "Altmetric - Mentions - Springer Nature - " + date + ".csv";
String xlsxFileAddress0 = downloadFilepath +   xlsxFileAddress;
File fa1 = new File(downloadFilepath + csvFileAddress);
for (int i1 = 2; i1 < 20; i1++) {
    	File directoryq = new File(downloadFilepath);
        File[] filesq = directoryq.listFiles();
        for (File fq : filesq){
        	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
        		String oldfilename = fq.getName(); 
            	fa1 = new File(downloadFilepath + oldfilename);
            	csvFileAddress = downloadFilepath + oldfilename; 
            	i1 = 250;
            }
        }
        Thread.sleep(10000);
}

if(fa1.exists()){
	int RowNum=-1;
		@SuppressWarnings("resource")
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("sheet1");
		String currentLine=null;
		RowNum=-1;
		BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
		while ((currentLine = br.readLine()) != null) {
			String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
			RowNum++;
			XSSFRow currentRow=sheet.createRow(RowNum);
			for(int i=0;i<str.length;i++){
				currentRow.createCell(i).setCellValue(str[i]);
			}
		}		
		FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress0);
		workBook.write(fileOutputStream);
		fileOutputStream.close();
		br.close();
		if(fa1.delete()){
		}

	File fa0 = new File(xlsxFileAddress0);
	if(fa0.exists()){
		Workbook wb = new Workbook();
		wb.loadFromFile(xlsxFileAddress0);
		wb.getCalculationMode();
		wb.calculateAllValue();
		wb.calculateFormulaValue(null);
		Worksheet sheet1 = wb.getWorksheets().get(0);  
		sheet1.getCellRange("AN2").setValue("News Storys");
		sheet1.getCellRange("AN3").setValue("Tweets");
		sheet1.getCellRange("AN4").setValue("Facebook posts");
		sheet1.getCellRange("AN5").setValue("Blog Posts");
		sheet1.getCellRange("AN6").setValue("Google+ posts");
		sheet1.getCellRange("AN7").setValue("Reddit + posts");
		sheet1.getCellRange("AN8").setValue("LinkedIn posts");
		sheet1.getCellRange("AN9").setValue("Videos");
		sheet1.getCellRange("AO2").setValue("News Story");
		sheet1.getCellRange("AO3").setValue("Tweet");
		sheet1.getCellRange("AO4").setValue("Facebook post");
		sheet1.getCellRange("AO5").setValue("Blog Post");
		sheet1.getCellRange("AO6").setValue("Google+ post");
		sheet1.getCellRange("AO7").setValue("Reddit + post");
		sheet1.getCellRange("AO8").setValue("LinkedIn post");
		sheet1.getCellRange("AO9").setValue("Video");
		sheet1.getCellRange("AM7").setValue("Reddit post");
		sheet1.getCellRange("AL7").setValue("Reddit posts");
		sheet1.getCellRange("AP2:AP9").setFormula("IFERROR(COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AO\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AN\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AM\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AL\"&ROW())),\"\")");
		sheet1.getCellRange("AP10").setFormula("COUNTA(A:A)-SUM(AP2:AP9)-1");
		sheet1.getCellRange("AP1").setFormula("ROUND(COUNTA(A:A)-1,\"0\")");
		sheet1.getCellRange("AP11").setFormula("SUM(AP2:AP10)");

		firstrows = 1;
		for (int i1 = 3; i1 < 100000; i1++) {
			if (!sheet1.getRange().get("A" + i1).getDisplayedText().contentEquals("")) 
			{firstrows= firstrows + 1;}else {
				i1 = 100000;
			}
		}
		
int firstrows1 = firstrows + 1;
		@SuppressWarnings("unused")
		SortColumn column = wb.getDataSorter().getSortColumns().add(7, SortComparsionType.Values, OrderBy.Descending);
		wb.getDataSorter().sort(sheet1.getCellRange("A1:AD" + firstrows1));
		int valickech = 0;	
		int valickechlast = 2;	
		int rowscheck = firstrows + 2;
		for (int i1 = 1; i1 < rowscheck; i1++) {
			if (!sheet1.getRange().get("H" + valickechlast).getDisplayedText().contentEquals(sheet1.getRange().get("H" + i1).getValue())) 
			{valickech = valickech + 1;} 
			valickechlast = valickechlast + 1; 
		}
		valickech = valickech - 1; 
		sheet1.getCellRange("AP12").setNumberValue(valickech);
		wb.getCalculationMode();
		wb.calculateAllValue();
		wb.calculateFormulaValue(null);
//		wb.saveToFile(downloadFilepath + "alt.xlsx", ExcelVersion.Version2013);
		
		for (int i1 = 2; i1 < 13; i1++) {
			String value = sheet1.getRange().get("AP" + i1).getDisplayedText();
			if (value.contentEquals("0")){
				sheet1.getRange().get("AP" + i1).setValue("");
			}
			}
		UpdateSlidePojo call1 = new UpdateSlidePojo(71,"<tca11>",sheet1.getRange().get("AP2").getDisplayedText());
		UpdateSlidePojo call2 = new UpdateSlidePojo(71,"<tca12>",sheet1.getRange().get("AP3").getDisplayedText());
		UpdateSlidePojo call3 = new UpdateSlidePojo(71,"<tca13>",sheet1.getRange().get("AP4").getDisplayedText());
		UpdateSlidePojo call4 = new UpdateSlidePojo(71,"<tca14>",sheet1.getRange().get("AP5").getDisplayedText());
		UpdateSlidePojo call5 = new UpdateSlidePojo(71,"<tca15>",sheet1.getRange().get("AP6").getDisplayedText());
		UpdateSlidePojo call6 = new UpdateSlidePojo(71,"<tca16>",sheet1.getRange().get("AP7").getDisplayedText());
		UpdateSlidePojo call7 = new UpdateSlidePojo(71,"<tca17>",sheet1.getRange().get("AP8").getDisplayedText());
		UpdateSlidePojo call8 = new UpdateSlidePojo(71,"<tca18>",sheet1.getRange().get("AP9").getDisplayedText());
		UpdateSlidePojo call9 = new UpdateSlidePojo(71,"<tca19>",sheet1.getRange().get("AP10").getDisplayedText());
		UpdateSlidePojo call10 = new UpdateSlidePojo(71,"<tca20>",sheet1.getRange().get("AP11").getDisplayedText());
		UpdateSlidePojo call11 = new UpdateSlidePojo(71,"<tca21>",sheet1.getRange().get("AP12").getDisplayedText());
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
		Thread.sleep(5000);
		checkvalues = checkvalues + "1";
	}
//} catch(Exception e1) {System.out.println("Exception");}
}}}}}
} catch(Exception e1) {System.out.println("Exception");}
}
}

try {
File dir1 = new File(downloadFilepath);
File[] file1 = dir1.listFiles();
for (File fq : file1){
	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
		String oldfilename = fq.getName(); 
    	File fa1 = new File(downloadFilepath + oldfilename);
    	fa1.delete();
    }
}
} catch(Exception e1) {}

//2020
for (int a = 1; a < 5; a++) {
File fa02 = new File(downloadFilepath + xlsxFileAddress1);
if(!fa02.exists()){
try {
driver.get("https://www.altmetric.com/explorer/mentions?mentioned_after=2020-01-01&mentioned_before=2020-12-31");
driver.manage().window().maximize();
Thread.sleep(3000);
if(!driver.findElements(By.xpath("//*[@type='email']")).isEmpty()){
	driver.findElement(By.xpath("//*[@type='email']")).sendKeys(altid);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(2500);
	driver.findElement(By.xpath("//*[@type='password']")).sendKeys(altp);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(5000);
}
String checkoutput1 = "";
String checkoutputresult1 = "";
if(!driver.findElements(By.className("edit-button")).isEmpty()){
	driver.findElement(By.className("edit-button")).click();
	Thread.sleep(1000);
	WebElement viewtext = driver.findElement(By.xpath("//*[@id='journal_id']"));
	JavascriptExecutor jsq = (JavascriptExecutor) driver;
	jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);
	driver.findElement(By.id("journal_id")).sendKeys(eISSN);
	Thread.sleep(4000);
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='AutocompleteFieldResults']")));
	Thread.sleep(100);
	checkoutput1 = driver.findElement(By.className("AutocompleteFieldResults")).getText();	
	if(checkoutput1.indexOf("o results found") > 0){
		checkoutput1 = "";
		checkoutputresult1 = "";
		driver.findElement(By.id("journal_id")).clear();
		Thread.sleep(1500);
		driver.findElement(By.id("journal_id")).sendKeys(pISSN);
		Thread.sleep(2500);
		checkoutput1 = driver.findElement(By.className("AutocompleteFieldResults")).getText();
	}
	if(checkoutput1.indexOf("o results found") < 0){
		Thread.sleep(1000);
	if(!driver.findElements(By.className("AutocompleteFieldResults")).isEmpty()){
		driver.findElement(By.className("AutocompleteFieldResults")).click();}
	if(!driver.findElements(By.className("AutocompleteField-item-list-item-remove-button")).isEmpty()){
		driver.findElement(By.className("execute")).click();
		Thread.sleep(5000);
		new WebDriverWait(driver, Duration.ofSeconds(200))
      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='summary']//span")));
		if(!driver.findElements(By.xpath("//*[@class='summary']//span")).isEmpty()){
			checkoutputresult1 = driver.findElement(By.xpath("//*[@class='summary']//span")).getText();}
		Thread.sleep(2500);
		if(checkoutputresult1.indexOf("no mentions") < 0){
		driver.findElement(By.xpath("//*[@class=\"export-option-selector\"]")).click();
		driver.findElement(By.partialLinkText("Download results as CSV (UTF-8)")).click();
		//driver.findElement(By.partialLinkText("Download results as CSV")).click();
		Thread.sleep(10000);

String csvFileAddress1 = downloadFilepath +  "Altmetric - Mentions - Springer Nature - " + date + ".csv";
String xlsxFileAddress10 = downloadFilepath +   xlsxFileAddress1;
File fa2 = new File(csvFileAddress1);
for (int i1 = 2; i1 < 20; i1++) {
	File directoryq = new File(downloadFilepath);
    File[] filesq = directoryq.listFiles();
    for (File fq : filesq){
    	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
    		String oldfilename = fq.getName(); 
    		fa2 = new File(downloadFilepath + oldfilename);
    		csvFileAddress1 = downloadFilepath + oldfilename; 
        	i1 = 250;
        }
    }	
    Thread.sleep(10000);
}

if(fa2.exists()){
	int RowNum=-1;      
		@SuppressWarnings("resource")
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet1 = workBook.createSheet("sheet1");
		String currentLine=null;
		RowNum = -1;
		BufferedReader br = new BufferedReader(new FileReader(csvFileAddress1));
		while ((currentLine = br.readLine()) != null) {
			String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
			RowNum++;
			XSSFRow currentRow=sheet1.createRow(RowNum);
			for(int i=0;i<str.length;i++){
				currentRow.createCell(i).setCellValue(str[i]);
			}
		}
		FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress10);
		workBook.write(fileOutputStream);
		fileOutputStream.close();
		br.close();
	if(fa2.delete()){
	}

		File fa2a = new File(xlsxFileAddress10);
		if(fa2a.exists()){
			Workbook wb1 = new Workbook();
			wb1.loadFromFile(xlsxFileAddress10);
			wb1.getCalculationMode();
			wb1.calculateAllValue();
			wb1.calculateFormulaValue(null);
			Worksheet sheet11 = wb1.getWorksheets().get(0); 
			sheet11.getCellRange("AN2").setValue("News Storys");
			sheet11.getCellRange("AN3").setValue("Tweets");
			sheet11.getCellRange("AN4").setValue("Facebook posts");
			sheet11.getCellRange("AN5").setValue("Blog Posts");
			sheet11.getCellRange("AN6").setValue("Google+ posts");
			sheet11.getCellRange("AN7").setValue("Reddit + posts");
			sheet11.getCellRange("AN8").setValue("LinkedIn posts");
			sheet11.getCellRange("AN9").setValue("Videos");
			sheet11.getCellRange("AO2").setValue("News Story");
			sheet11.getCellRange("AO3").setValue("Tweet");
			sheet11.getCellRange("AO4").setValue("Facebook post");
			sheet11.getCellRange("AO5").setValue("Blog Post");
			sheet11.getCellRange("AO6").setValue("Google+ post");
			sheet11.getCellRange("AO7").setValue("Reddit + post");
			sheet11.getCellRange("AO8").setValue("LinkedIn post");
			sheet11.getCellRange("AO9").setValue("Video");
			sheet11.getCellRange("AM7").setValue("Reddit post");
			sheet11.getCellRange("AL7").setValue("Reddit posts");
			sheet11.getCellRange("AP2:AP9").setFormula("IFERROR(COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AO\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AN\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AM\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AL\"&ROW())),\"\")");
			sheet11.getCellRange("AP10").setFormula("COUNTA(A:A)-SUM(AP2:AP9)-1");
			sheet11.getCellRange("AP11").setFormula("SUM(AP2:AP10)");
			secondrows = 1;
			for (int i1 = 3; i1 < 100000; i1++) {
				if (!sheet11.getRange().get("A" + i1).getDisplayedText().contentEquals("")) 
				{secondrows = secondrows + 1;}else {
					i1 = 100000;
				}
			}
			int secondrows1 = secondrows + 1;
			@SuppressWarnings("unused")
			SortColumn column = wb1.getDataSorter().getSortColumns().add(7, SortComparsionType.Values, OrderBy.Descending);
			wb1.getDataSorter().sort(sheet11.getCellRange("A1:AD" + secondrows1));
			int valickech = 0;	
			int valickechlast = 2;	
			int rowscheck2 = secondrows + 2;
			for (int i1 = 1; i1 < rowscheck2; i1++) {
				if (!sheet11.getRange().get("H" + valickechlast).getDisplayedText().contentEquals(sheet11.getRange().get("H" + i1).getValue())) 
				{valickech = valickech + 1;}
				valickechlast = valickechlast + 1; 
			}
			valickech = valickech - 1;
			sheet11.getCellRange("AP12").setNumberValue(valickech);
			wb1.getCalculationMode();
			wb1.calculateAllValue();
			wb1.calculateFormulaValue(null);
			for (int i1 = 2; i1 < 13; i1++) {
				String value = sheet11.getRange().get("AP" + i1).getDisplayedText();
				if (value.contentEquals("0")){
					sheet11.getRange().get("AP" + i1).setValue("");
				}
				}
			UpdateSlidePojo callb1 = new UpdateSlidePojo(71,"<tca22>",sheet11.getRange().get("AP2").getDisplayedText());
			UpdateSlidePojo callb2 = new UpdateSlidePojo(71,"<tca23>",sheet11.getRange().get("AP3").getDisplayedText());
			UpdateSlidePojo callb3 = new UpdateSlidePojo(71,"<tca24>",sheet11.getRange().get("AP4").getDisplayedText());
			UpdateSlidePojo callb4 = new UpdateSlidePojo(71,"<tca25>",sheet11.getRange().get("AP5").getDisplayedText());
			UpdateSlidePojo callb5 = new UpdateSlidePojo(71,"<tca26>",sheet11.getRange().get("AP6").getDisplayedText());
			UpdateSlidePojo callb6 = new UpdateSlidePojo(71,"<tca27>",sheet11.getRange().get("AP7").getDisplayedText());
			UpdateSlidePojo callb7 = new UpdateSlidePojo(71,"<tca28>",sheet11.getRange().get("AP8").getDisplayedText());
			UpdateSlidePojo callb8 = new UpdateSlidePojo(71,"<tca29>",sheet11.getRange().get("AP9").getDisplayedText());
			UpdateSlidePojo callb9 = new UpdateSlidePojo(71,"<tca30>",sheet11.getRange().get("AP10").getDisplayedText());
			UpdateSlidePojo callb10 = new UpdateSlidePojo(71,"<tca31>",sheet11.getRange().get("AP11").getDisplayedText());
			UpdateSlidePojo callb11 = new UpdateSlidePojo(71,"<tca32>",sheet11.getRange().get("AP12").getDisplayedText());
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
			Thread.sleep(5000);
			checkvalues = checkvalues + "1";
		}
//		} catch(Exception e1) {System.out.println("Exception");}

}}}}}
} catch(Exception e1) {System.out.println("Exception");}
}
}
try {
File dir1 = new File(downloadFilepath);
File[] file1 = dir1.listFiles();
for (File fq : file1){
	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
		String oldfilename = fq.getName(); 
    	File fa1 = new File(downloadFilepath + oldfilename);
    	fa1.delete();
    }
}
} catch(Exception e1) {}
//2021 
for (int a = 1; a < 5; a++) {
File fa03 = new File(downloadFilepath + xlsxFileAddress11);
if(!fa03.exists()){
try {
	String checkoutput11 = "";
	String checkoutputresult11 = "";
driver.get("https://www.altmetric.com/explorer/mentions?mentioned_after=2021-01-01&mentioned_before=2021-12-31");
driver.manage().window().maximize();
Thread.sleep(3000);
if(!driver.findElements(By.xpath("//*[@type='email']")).isEmpty()){
	driver.findElement(By.xpath("//*[@type='email']")).sendKeys(altid);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(2500);
	driver.findElement(By.xpath("//*[@type='password']")).sendKeys(altp);
	driver.findElement(By.xpath("//*[@name='commit']")).click();
	Thread.sleep(5000);
}
if(!driver.findElements(By.className("edit-button")).isEmpty()){
	driver.findElement(By.className("edit-button")).click();
	Thread.sleep(1000);
	WebElement viewtext = driver.findElement(By.xpath("//*[@id='journal_id']"));
	JavascriptExecutor jsq = (JavascriptExecutor) driver;
	jsq.executeScript("arguments[0].scrollIntoView(true);", viewtext);
	driver.findElement(By.id("journal_id")).sendKeys(eISSN);
	Thread.sleep(4000);
	new WebDriverWait(driver, Duration.ofSeconds(10))
	.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='AutocompleteFieldResults']")));
	Thread.sleep(100);
	checkoutput11 = driver.findElement(By.className("AutocompleteFieldResults")).getText();	
	if(checkoutput11.indexOf("o results found") > 0){
		driver.findElement(By.id("journal_id")).clear();
		Thread.sleep(2500);
		driver.findElement(By.id("journal_id")).sendKeys(pISSN);
		Thread.sleep(2500);
		checkoutput11 = "";
		checkoutput11 = driver.findElement(By.className("AutocompleteFieldResults")).getText();	
	}
	if(checkoutput11.indexOf("o results found") < 0){
		driver.findElement(By.className("AutocompleteFieldResults")).click();
	if(!driver.findElements(By.className("AutocompleteField-item-list-item-remove-button")).isEmpty()){
		driver.findElement(By.className("execute")).click();
		Thread.sleep(5000);
		new WebDriverWait(driver, Duration.ofSeconds(200))
      .until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@class='summary']//span")));
		if(!driver.findElements(By.xpath("//*[@class='summary']//span")).isEmpty()){
		checkoutputresult11 = driver.findElement(By.xpath("//*[@class='summary']//span")).getText();}
		Thread.sleep(2500);
		if(checkoutputresult11.indexOf("no mentions") < 0){
		driver.findElement(By.xpath("//*[@class=\"export-option-selector\"]")).click();
		driver.findElement(By.partialLinkText("Download results as CSV (UTF-8)")).click();
		//driver.findElement(By.partialLinkText("Download results as CSV")).click();
		Thread.sleep(10000);
		
//2021 
String csvFileAddress11 = downloadFilepath +  "Altmetric - Mentions - Springer Nature - " + date + ".csv";
String xlsxFileAddress110 = downloadFilepath +   xlsxFileAddress11;
File fa3 = new File(csvFileAddress11);

for (int i1 = 2; i1 < 20; i1++) {
	File directoryq = new File(downloadFilepath);
    File[] filesq = directoryq.listFiles();
    for (File fq : filesq){
    	if (fq.getName().startsWith("Altmetric - Mentions - Springer Nature")){
    		String oldfilename = fq.getName(); 
    		fa3 = new File(downloadFilepath + oldfilename);
    		csvFileAddress11 = downloadFilepath + oldfilename; 
        	i1 = 250;
        }
    }	
    Thread.sleep(10000);
}

if(fa3.exists()){
		@SuppressWarnings("resource")
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet11 = workBook.createSheet("sheet1");
		String currentLine=null;
		int RowNum=-1;
		BufferedReader br = new BufferedReader(new FileReader(csvFileAddress11));
		while ((currentLine = br.readLine()) != null) {
			String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
			RowNum++;
			XSSFRow currentRow=sheet11.createRow(RowNum);
			for(int i=0;i<str.length;i++){
				currentRow.createCell(i).setCellValue(str[i]);
			}
		}
		FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress110);
		workBook.write(fileOutputStream);
		fileOutputStream.close();
		br.close();
		if(fa3.delete()){
		}

		File fa4 = new File(xlsxFileAddress110);
		if(fa4.exists()){
			Workbook wb11 = new Workbook();
			wb11.loadFromFile(xlsxFileAddress110);
			wb11.getCalculationMode();
			wb11.calculateAllValue();
			wb11.calculateFormulaValue(null);
			Worksheet sheet111 = wb11.getWorksheets().get(0);

			sheet111.getCellRange("AG2").setFormula("COUNTA(A:A)");		
			int tstring1 = 2;
					for (int j = 2; j <= 100000; j++) {
				String valuecheck = sheet111.getRange().get("P" + j).getDisplayedText();
				if(valuecheck.isEmpty()) {
					j = 100000;
					tstring1 = tstring1 - 1;
				}else {
					sheet111.getRange().get("P" + j).setValue(valuecheck);
					tstring1 = tstring1 + 1; 
				}
			}
			@SuppressWarnings("unused")
			SortColumn column = wb11.getDataSorter().getSortColumns().add(15, SortComparsionType.Values, OrderBy.Descending);
			wb11.getDataSorter().sort(sheet111.getCellRange("A1:AD" + tstring1));
			
			sheet111.getCellRange("AN2").setValue("News Storys");
			sheet111.getCellRange("AN3").setValue("Tweets");
			sheet111.getCellRange("AN4").setValue("Facebook posts");
			sheet111.getCellRange("AN5").setValue("Blog Posts");
			sheet111.getCellRange("AN6").setValue("Google+ posts");
			sheet111.getCellRange("AN7").setValue("Reddit + posts");
			sheet111.getCellRange("AN8").setValue("LinkedIn posts");
			sheet111.getCellRange("AN9").setValue("Videos");
			sheet111.getCellRange("AO2").setValue("News Story");
			sheet111.getCellRange("AO3").setValue("Tweet");
			sheet111.getCellRange("AO4").setValue("Facebook post");
			sheet111.getCellRange("AO5").setValue("Blog Post");
			sheet111.getCellRange("AO6").setValue("Google+ post");
			sheet111.getCellRange("AO7").setValue("Reddit + post");
			sheet111.getCellRange("AO8").setValue("LinkedIn post");
			sheet111.getCellRange("AO9").setValue("Video");
			sheet111.getCellRange("AM7").setValue("Reddit post");
			sheet111.getCellRange("AL7").setValue("Reddit posts");
			sheet111.getCellRange("AP2:AP9").setFormula("IFERROR(COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AO\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AN\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AM\"&ROW()))+COUNTIF(INDIRECT(\"A:A\"),INDIRECT(\"AL\"&ROW())),\"\")");
			sheet111.getCellRange("AP10").setFormula("COUNTA(A:A)-SUM(AP2:AP9)-1");
			sheet111.getCellRange("AP11").setFormula("SUM(AP2:AP10)");
//			wb11.saveToFile(downloadFilepath + "//alt.xlsx", ExcelVersion.Version2013);
			wb11.getCalculationMode();
			wb11.calculateAllValue();
			wb11.calculateFormulaValue(null);
			WebDriver driver1 = new ChromeDriver(options);
			String res2a = "";
			String res2b = "";
			String res2c = "";
			String res2d = "";
//			int i14 = 1;
			int totalunique = 0;
			String one = "";
			String two = "";
			int valuecheck1 = 2;
			System.out.println(valuecheck1);
			for (int j = 1; j <= tstring1; j++) {
		one = one + sheet111.getRange().get("H" + j).getDisplayedText();
		two = sheet111.getRange().get("H" + valuecheck1).getDisplayedText();
		if(one.indexOf(two) < 0) {	
			if (totalunique < 10) {
			String res1a = sheet111.getRange().get("P" + valuecheck1).getDisplayedText();
			String res1b = sheet111.getRange().get("H" + valuecheck1).getDisplayedText();
			String res1c = sheet111.getRange().get("Q" + valuecheck1).getDisplayedText();
			String res1d = sheet111.getRange().get("O" + valuecheck1).getDisplayedText();
			res2a = res2a + '\n' + res1a;
			res2b = res2b + '\n' + res1b;
			try {
			driver1.get(res1c);			
			Thread.sleep(300);
			String author1 = "";
			if(!driver1.findElements(By.xpath("(//*[@class='truncated'])")).isEmpty()){
				author1 = driver1.findElement(By.xpath("(//*[@class='truncated'])")).getText();
				Thread.sleep(100);
			}else {
			driver1.close();
			driver1.quit();
			}		
			res2c = res2c + '\n' + author1;
			res2d = res2d + '\n' + res1d;
			} catch(Exception e1) {j = j + tstring1;}
//			i14++;
			}
			totalunique = totalunique + 1; 
		}else {} 
		valuecheck1++;
	}
			try {
				driver1.close();
				driver1.quit();
				} catch(Exception e1) {}
			
	resa = res2a;
	resb = res2b;
	resc = res2c;
	resd = res2d;
	sheet111.getCellRange("AP12").setNumberValue(totalunique);

	for (int i1 = 2; i1 < 13; i1++) {
	String value = sheet111.getRange().get("AP" + i1).getDisplayedText();
	if (value.contentEquals("0")){
		sheet111.getRange().get("AP" + i1).setValue("");
	}
	}
	UpdateSlidePojo callc1 = new UpdateSlidePojo(71,"<tca33>",sheet111.getRange().get("AP2").getDisplayedText());
	UpdateSlidePojo callc2 = new UpdateSlidePojo(71,"<tca34>",sheet111.getRange().get("AP3").getDisplayedText());
	UpdateSlidePojo callc3 = new UpdateSlidePojo(71,"<tca35>",sheet111.getRange().get("AP4").getDisplayedText());
	UpdateSlidePojo callc4 = new UpdateSlidePojo(71,"<tca36>",sheet111.getRange().get("AP5").getDisplayedText());
	UpdateSlidePojo callc5 = new UpdateSlidePojo(71,"<tca37>",sheet111.getRange().get("AP6").getDisplayedText());
	UpdateSlidePojo callc6 = new UpdateSlidePojo(71,"<tca38>",sheet111.getRange().get("AP7").getDisplayedText());
	UpdateSlidePojo callc7 = new UpdateSlidePojo(71,"<tca39>",sheet111.getRange().get("AP8").getDisplayedText());
	UpdateSlidePojo callc8 = new UpdateSlidePojo(71,"<tca40>",sheet111.getRange().get("AP9").getDisplayedText());
	UpdateSlidePojo callc9 = new UpdateSlidePojo(71,"<tca41>",sheet111.getRange().get("AP10").getDisplayedText());
	UpdateSlidePojo callc10 = new UpdateSlidePojo(71,"<tca42>",sheet111.getRange().get("AP11").getDisplayedText());
	UpdateSlidePojo callc11 = new UpdateSlidePojo(71,"<tca43>",sheet111.getRange().get("AP12").getDisplayedText());
	UpdateSlidePojo callc12 = new UpdateSlidePojo(11,"<tca42>",sheet111.getRange().get("AP11").getDisplayedText());
		updateSlides.add(callc1);
		updateSlides.add(callc2);
		updateSlides.add(callc3);
		updateSlides.add(callc4);
		updateSlides.add(callc5);
		updateSlides.add(callc6);
		updateSlides.add(callc7);
		updateSlides.add(callc8);
		updateSlides.add(callc9);
		updateSlides.add(callc10);
		updateSlides.add(callc11);
		updateSlides.add(callc12);
		Thread.sleep(5000);
		checkvalues = checkvalues + "1";
		}
		if (!checkvalues.contentEquals("")) {
		pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides);
		sheet0.updateCellvalue(sheetsService, sheetname + "!M35", "1", sheetid);
		}
		if (!checkvalues.contentEquals("")) {
	sheet0.updateCellvalue(sheetsService, sheetname + "!M36", resa + "#" + resb + "#" + resc + "#" + resd,sheetid); 
		}

}
}}
}}
} catch(Exception e1) {System.out.println("Exception");}
}
}

try {
	driver.quit();
	driver1.quit();
} catch(Exception e1) {System.out.println("Exceptionbex");}
}
}
}
}

  @AfterTest
  public void afterTest() {	 
try {
		driver.quit();
		driver1.quit();
		} catch(Exception e1) {System.out.println("Exceptionbex");}
  }

}