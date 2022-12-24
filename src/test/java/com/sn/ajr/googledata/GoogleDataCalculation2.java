package com.sn.ajr.googledata; 
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException; 
import java.util.List; 
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets;
import com.sn.ajr.ProjGSheet;  
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
public class GoogleDataCalculation2 {
	  @Test
	public static void classname() throws InterruptedException, IOException, GeneralSecurityException { 
		System.out.println("WOS calculation start");
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
//		String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheet1 = new ProjGSheet();
		Sheets sheetsService1 = sheet1.getSheetsService("Google Sheets Example"); 
    	String sheetname = "";
    	try {
        	List<List<Object>> values = sheet1.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
    		for (List<?> row : values) {
        	sheetname = (String) row.get(0); 
    		}
        	} catch (Exception e1) {}
    		
        	if(!sheetname.contentEquals("")) {
        		String sheetid = "";
            	try {
            	List<List<Object>> values = sheet1.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
            	for (List<?> row : values) {
            	sheetid = (String) row.get(0);
            	}
            	 } catch (Exception e1) {}
            	if(!sheetid.contentEquals("")) { 
//		File currentDirFile = new File(".");
//		String rootpath1 = currentDirFile.getAbsolutePath();
//		String rootpath = rootpath1.replace(".", "");
//		String downloadFilepath = rootpath + "DownloadingTempfiles\\"; 
		ProjGSheet sheet = new ProjGSheet();
		 Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
		 
//Country usage4
		 String csvFilecountry = downloadFilepath +  "Country.csv";
			String xlsxFilecountry = downloadFilepath +   "Country.xlsx";
			File fac = new File(csvFilecountry);
			for (int i1 = 2; i1 < 150; i1++) {
				Thread.sleep(1000);
				if(fac.exists()){
					i1 = 150;
				}	
			}
			if(fac.exists()){
				int RowNum=-1;
				try {
					@SuppressWarnings("resource")
					XSSFWorkbook workBook = new XSSFWorkbook();
					XSSFSheet sheetc = workBook.createSheet("sheet1");
					String currentLine=null;
					RowNum=-1;
					BufferedReader br = new BufferedReader(new FileReader(csvFilecountry));
					while ((currentLine = br.readLine()) != null) {
						String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
						RowNum++;
						XSSFRow currentRow=sheetc.createRow(RowNum);
						for(int i=0;i<str.length;i++){
							currentRow.createCell(i).setCellValue(str[i]);
						}
					}
					FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFilecountry);
					workBook.write(fileOutputStream);
					fileOutputStream.close();
					br.close();
					System.out.println("Done");
				} catch (Exception ex) {System.out.println(ex.getMessage()+"Exception in try");}
		 
try{
String inputsheet = downloadFilepath + "Country.xlsx"; 

Workbook wbb1 = new Workbook();
wbb1.loadFromFile(inputsheet);
Worksheet sheetb1 = wbb1.getWorksheets().get(0); 
String medvalue = "^";
String countryname = "";
String countryvalue = "";
for (int i12 = 2; i12 < 1000; i12++) {
	String res1a = sheetb1.getRange().get("A" + i12).getDisplayedText();
	String res1b = sheetb1.getRange().get("B" + i12).getDisplayedText();
	res1b = res1b.replace("\"", "");
	if (countryname.equals("")) {
		medvalue = "";
		}else {
		medvalue= "^";
		}
	countryname = countryname + medvalue + res1a;
	countryvalue = countryvalue + medvalue + res1b;
	int icheck = i12 + 1;
	String res1c = sheetb1.getRange().get("B" + icheck).getDisplayedText();
	if(res1c.contentEquals("")){
		i12 = 1000;
	}
	res1a = "";
	res1b = "";
}
Thread.sleep(500);
sheet.updateCellvalue(sheetsService, sheetname + "!AD1", countryname,sheetid);
sheet.updateCellvalue(sheetsService, sheetname + "!AE1", countryvalue,sheetid);
System.out.println("4.5 Done");
} catch(Exception e1) {System.out.println("excelptioninca");} }
	}}
	}}