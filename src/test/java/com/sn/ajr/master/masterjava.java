package com.sn.ajr.master;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.beust.jcommander.internal.Lists;
import com.google.api.services.sheets.v4.Sheets;
import com.sn.ajr.ProjGSheet;
import com.sn.ajr.altmetric.AltmetricMaster;
import com.sn.ajr.citation.Awos;
import com.sn.ajr.citation.Bconversionofexcel;
import com.sn.ajr.downloads.part1;
import com.sn.ajr.citation.Ccalculation;
import com.sn.ajr.downloads.part2;
import com.sn.ajr.downloads.sjr;
import com.sn.ajr.googledata.GoogleData;
import com.sn.ajr.googledata.GoogleDataCalculation1;
import com.sn.ajr.googledata.GoogleDataCalculation2;
import com.sn.ajr.jcr.JCR;
import com.sn.ajr.jcr.ProjJCRcalculation_CitingCited;
import com.sn.ajr.jcr.ProjJcrCategory1;
import com.sn.ajr.jcr.ProjJcrCategory2;
import com.sn.ajr.jcr.ProjJcrCategory3;
import com.sn.ajr.jcr.ProjJcrCategory4;
import com.sn.ajr.jcr.ProjJcrCategory5;
import com.sn.ajr.jcr.ProjJcrCategory6;
import com.sn.ajr.screenshots.Uploadscreenshots;
public class masterjava {
	public static WebDriver driver1;
    String request = "NoRequests";
    String sheetidq = "";
    File currentDirFile = new File(".");
	String rootpath1 = currentDirFile.getAbsolutePath();
	String rootpath = rootpath1.replace(".", "");
	String downloadFilepath = rootpath + "DownloadingTempfiles\\";
	String livepath = rootpath + "batchfiles\\live.txt"; 
//	 @BeforeTest
//		public void beforeTest() throws InterruptedException, IOException, GeneralSecurityException {
    @BeforeTest
	public void beforeTest() throws InterruptedException, IOException, GeneralSecurityException {
		String sheetidq = "";
		String createreq = "";
    	File f0 = new File(livepath);
    	if(f0.exists()){
    		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
    		createreq = "https://script.google.com/macros/s/AKfycbwrTqkQezj2f9o8G9BJda8R6BrgUggj-TBqhtR_3xJnisio4hC9CbewJTE2SDKV9mA/exec";
    	}else {
    		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
    		createreq = "https://script.google.com/macros/s/AKfycbwpsbUZgdEs2FgdsjD9_js6H1FS9UB9CkWznOKYeOGy5QTG_NOYCpqzWrqXY1uNo6U/exec";
    	}
		ProjGSheet sheet = new ProjGSheet();
		 Sheets sheetsService = sheet.getSheetsService("Google Sheets Example"); 
		 for (int i1 = 2; i1 < 999999999;) {
		try {
		List<List<Object>> valuesaa = sheet.getCellData(sheetsService, "queuecheck!B2", sheetidq);
    	for (List<?> row : valuesaa) {request = (String) row.get(0);}
		 } catch(Exception e1) {}		
		if(!request.contentEquals("NoRequests")) {
			i1 = 999999999;
		}else {Thread.sleep(30000);}
		}		 
    	if(!request.contentEquals("NoRequests")) {
        	try {
    		File f = new File(downloadFilepath);
    		FileUtils.forceMkdir(f);
    		FileUtils.cleanDirectory(f);
    		} catch (IOException e) {e.printStackTrace();}    	
    	System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
		WebDriver driver = new ChromeDriver(); 
	    driver.get(createreq);
	    Thread.sleep(1000);
		request = (String) driver.findElement(By.xpath("//body")).getText();
	    driver.close();
    	}
    }
   @Test
public void part1() throws InterruptedException, IOException, GeneralSecurityException {
	   if(!request.contentEquals("NoRequests")) {
		   String matadatareq = "";
	    	File f0 = new File(livepath);
	    	if(f0.exists()){
	    		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
	    		matadatareq = "https://script.google.com/macros/s/AKfycbwEldD_SYQ8CsDlk32tmn-CYlzqnosq7WFdfZdJWcRIo09unP6nSwSNOB4LdEXeDo45xg/exec";
	    	}else {
	    		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
	    		matadatareq = "https://script.google.com/macros/s/AKfycby1jSTbsJB1Vmjs_zPypnSnvG6QVlyJwYwUJOqgmW7h-9VCCTVIEX35HS2N_eMfUfGV3g/exec";
	    	}
		System.setProperty("webdriver.gecko.driver", "./geckodriver/geckodriver.exe");
   		driver1 = new FirefoxDriver();
   		//System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
   		//WebDriver driver1 = new ChromeDriver(); 
   		driver1.get(matadatareq);
   		Thread.sleep(600000);
   		driver1.close();
	   }
}
    @Test
public void part2() throws InterruptedException, IOException, GeneralSecurityException {
    	if(!request.contentEquals("NoRequests")) {
    		part1.classname();
    		sjr.classname();
       	   	sjr.classname(); 
    		GoogleData.classname();
    		GoogleData.classname();
    		GoogleDataCalculation1.classname();
    		GoogleDataCalculation2.classname();
    		part2.mainclass(); 
    		part2.mainclass(); 
    		AltmetricMaster.classname(); 
	        Awos.classname(); 
	 		Bconversionofexcel.classname();
	 		Ccalculation.classname();
	 		JCR.classname();
			ProjJCRcalculation_CitingCited.mainclass(); 
			ProjJcrCategory1.mainclass();
			ProjJcrCategory2.mainclass();
			ProjJcrCategory3.mainclass();
			ProjJcrCategory4.mainclass();
			ProjJcrCategory5.mainclass();
			ProjJcrCategory6.mainclass();	
			Uploadscreenshots.uploadfilebyid(null);
       }
}
    @Test(dependsOnMethods="part2")
    public void recheck() throws InterruptedException, IOException, GeneralSecurityException {
    	   if(!request.contentEquals("NoRequests")) {
    		}
    }

  @AfterTest
  public void afterTest() throws IOException, GeneralSecurityException, InterruptedException {
    		if(!request.contentEquals("NoRequests")) {
    	 	  	String livepath = rootpath + "batchfiles\\live.txt"; 
        		String sheetidq = "";
        		String downloadreq = "";
        		File f0 = new File(livepath);
        		if(f0.exists()){
          		sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";	
          		downloadreq = "https://script.google.com/macros/s/AKfycbxBOQqtZcl3vDvACyOmF8FzS4MLEiNy0s0r_wpWo7ywPYBpZB771DRuq_2RDnG8USEpYQ/exec";
        		}else {
          		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY";
          		downloadreq = "https://script.google.com/macros/s/AKfycbz9dVkWTYuIeKInGzTePCSGu5K6OOrxHvIgPt6ISr9rFneJ8sVnkle00PSr4XsyiEg20w/exec";
        		}        		
    		System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");
    		WebDriver driver = new ChromeDriver(); 
    	    driver.get(downloadreq);
    	    
//    	    ProjGSheet sheet = new ProjGSheet();
//    		Sheets sheetsService1 = sheet.getSheetsService("Google Sheets Example"); 
        	ProjGSheet sheet1 = new ProjGSheet();
    		Sheets sheetsService = sheet1.getSheetsService("Google Sheets Example");
    		
    		for (int i1 = 1; i1 < 100; i1++) {
    		String idcheck1 = "";
    		try {
    		List<List<Object>> valuesaa = sheet1.getCellData(sheetsService, "queuecheck!B2", sheetidq);
        	for (List<?> row : valuesaa) {idcheck1 = (String) row.get(0);}
        	} catch(Exception e1) {}
    		try {
    		if(!idcheck1.contentEquals("")) {
        		List<List<Object>> valuesaa = sheet1.getCellData(sheetsService, "queuecheck!B2", sheetidq);
            	for (List<?> row : valuesaa) {idcheck1 = (String) row.get(0);}
    		}
    		} catch(Exception e1) {}
    		try {
        		if(!idcheck1.contentEquals(request)) {
            		List<List<Object>> valuesaa = sheet1.getCellData(sheetsService, "queuecheck!B2", sheetidq);
                	for (List<?> row : valuesaa) {idcheck1 = (String) row.get(0);}
        		}
        	} catch(Exception e1) {}
    		try {
    		if(!idcheck1.contentEquals(request)) {
    			if(!idcheck1.contentEquals("")) {
    			System.out.println("Ending This Rrequest because of " + idcheck1);
    				i1 = 100;	
    			}
    			
    		}
    		} catch(Exception e1) {}
    			Thread.sleep(20000);     
    			System.out.print("attempt " + i1);
    		if(idcheck1.contentEquals("21")) {
    			try {
            		List<List<Object>> valuesaa = sheet1.getCellData(sheetsService, "queuecheck!A2", sheetidq);
                	for (List<?> row : valuesaa) {idcheck1 = (String) row.get(0);}
                	} catch(Exception e2) {}
    			//sheet.updateCellvalue(sheetsService1,"queue!H" + idcheck1, "Error",sheetidq);
    			}
        	}
    		try {
    		driver.quit(); 
    		driver1.quit();
    		} catch(Exception e1) {}
    		}
    		try {
    		ProcessBuilder pb = new ProcessBuilder("cmd", "/c", "cleanup.bat");
        	File dir = new File(rootpath + "batchfiles");
        	pb.directory(dir);
        	Process p = pb.start();
    		} catch(Exception e1) {}
    		try {
    		File currentDirFile = new File(".");
    		String rootpath1 = currentDirFile.getAbsolutePath();
    		String rootpath = rootpath1.replace(".", "");
    		String downloadFilepath = rootpath + "\\";
    		System.out.println("cheking next request");
    		TestListenerAdapter tla = new TestListenerAdapter();
    		TestNG testng = new TestNG();
    		List<String> suites = Lists.newArrayList();
    		suites.add(downloadFilepath + "testng.xml");
//    		suites.add(downloadFilepath + "testng2.xml");
    		testng.setTestSuites(suites);
    		testng.run();
    		} catch(Exception e1) {}
    	}

		 
}
