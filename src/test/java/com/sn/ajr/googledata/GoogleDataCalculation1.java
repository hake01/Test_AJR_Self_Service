package com.sn.ajr.googledata;
import org.testng.annotations.Test;
import com.google.api.services.sheets.v4.Sheets; 
import com.sn.ajr.ProjGSheet; 
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.testng.annotations.BeforeTest;
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
import org.testng.annotations.AfterTest;

public class GoogleDataCalculation1 {
//	spire.xls
  @Test
  public static void classname() throws InterruptedException, IOException, GeneralSecurityException {
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
//	  String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example"); 
		String sheetname = ""; 
  	try {
      	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
  		for (List<?> row : values) { 
  			sheetname = (String) row.get(0); 
  		}
      	} catch (Exception e1) {}
  		
      	if(!sheetname.contentEquals("")) {
      		String sheetid = "";
        	try {
        	List<List<Object>> values = sheetq.getCellData(sheetsService1, "Details!C10:C10",sheetidq); 
        	for (List<?> row : values) {
        	sheetid = (String) row.get(0);
        	}
        	 } catch (Exception e1) {}
        	if(!sheetid.contentEquals("")) {
		ProjGSheet sheet = new ProjGSheet();
		Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
  	
//    	File currentDirFile = new File(".");
//		String rootpath1 = currentDirFile.getAbsolutePath();
//		String rootpath = rootpath1.replace(".", "");
//		String downloadFilepath = rootpath + "DownloadingTempfiles\\";
  	
//		rejection
		String csvFile3 = downloadFilepath +  "Rejection_Tracker.csv";
		String xlsxFile3 = downloadFilepath +   "Rejection_Tracker.xlsx";
		File fa2 = new File(csvFile3);
		for (int i1 = 2; i1 < 150; i1++) {
			Thread.sleep(1000);
			if(fa2.exists()){
				i1 = 150;
			}	
		}
		if(fa2.exists()){
			int RowNum=-1;
			try {
				@SuppressWarnings("resource")
				XSSFWorkbook workBook = new XSSFWorkbook();
				XSSFSheet sheet1 = workBook.createSheet("sheet1");
				String currentLine=null;
				RowNum=-1;
				BufferedReader br = new BufferedReader(new FileReader(csvFile3));
				while ((currentLine = br.readLine()) != null) {
					String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
					RowNum++;
					XSSFRow currentRow=sheet1.createRow(RowNum);
					for(int i=0;i<str.length;i++){
						currentRow.createCell(i).setCellValue(str[i]);
					}
				}
				FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFile3);
				workBook.write(fileOutputStream);
				fileOutputStream.close();
				br.close();
				System.out.println("Done");
			} catch (Exception ex) {System.out.println(ex.getMessage()+"Exception in try");}
	try {
  	Workbook wba1 = new Workbook();
      wba1.loadFromFile(xlsxFile3);
      Worksheet sheet1 = wba1.getWorksheets().get(0);        
      sheet1.getCellRange("I2:I1000").setFormula("=IFERROR(VALUE(SUBSTITUTE(INDIRECT(\"E\"&ROW()),CHAR(34),\"\")),\"\")");
      sheet1.getCellRange("J1:J1").setFormula("=SUMIF(B:B,\"Springer Nature\",I:I)");
      sheet1.getCellRange("K1:K1").setFormula("=SUM(I:I)-J1");
      sheet1.getCellRange("J2:J2").setFormula("=SUBSTITUTE(A2&\"^\"&A3&\"^\"&A4&\"^\"&A5&\"^\"&A6&\"^\"&A7&\"^\"&A8&\"^\"&A9&\"^\"&A10&\"^\"&A11&\"^\"&A12&\"^\"&A13&\"^\"&A14&\"^\"&A15&\"^\"&A16&\"^\"&A17&\"^\"&A18&\"^\"&A19&\"^\"&A20&\"^\"&A21,CHAR(34),\"\")");
      sheet1.getCellRange("K2:K2").setFormula("=SUBSTITUTE(I2&\"^\"&I3&\"^\"&I4&\"^\"&I5&\"^\"&I6&\"^\"&I7&\"^\"&I8&\"^\"&I9&\"^\"&I10&\"^\"&I11&\"^\"&I12&\"^\"&I13&\"^\"&I14&\"^\"&I15&\"^\"&I16&\"^\"&I17&\"^\"&I18&\"^\"&I19&\"^\"&I20&\"^\"&I21,CHAR(34),\"\")");
      wba1.getCalculationMode();
      wba1.calculateAllValue();
      wba1.calculateFormulaValue(null);
//      wba1.saveToFile(downloadFilepath + "//rejection.xlsx", ExcelVersion.Version2013);
      	String aa1 = sheet1.getRange().get("J1").getDisplayedText();
      	String ab1 = sheet1.getRange().get("K1").getDisplayedText();
      	String ac1 = sheet1.getRange().get("J2").getDisplayedText();
      	String ad1 = sheet1.getRange().get("K2").getDisplayedText();
      	System.out.println(aa1);
      	System.out.println(ab1);
      	sheet.updateCellvalue(sheetsService, sheetname + "!B11", aa1,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!C11", ab1,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A61", ac1,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A62", ad1,sheetid);   
      
  	} catch(Exception e1) {}  	 
	}

//		4.2
		String csvFileAddress1 = downloadFilepath +  "4.2.csv"; //csv file address
		String xlsxFileAddress1 = downloadFilepath +   "4.2.xlsx"; //xlsx file address
		File fa11 = new File(csvFileAddress1);
		for (int i1 = 2; i1 < 150; i1++) {
			Thread.sleep(1000);
			if(fa11.exists()){
				i1 = 150;
			}	
		}

		if(fa11.exists()){
			int RowNum=-1;
			try {
				@SuppressWarnings("resource")
				XSSFWorkbook workBook = new XSSFWorkbook();
				XSSFSheet sheet1 = workBook.createSheet("sheet1");
				String currentLine=null;
				RowNum=-1;
				BufferedReader br = new BufferedReader(new FileReader(csvFileAddress1));
				while ((currentLine = br.readLine()) != null) {
					String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
					RowNum++;
					XSSFRow currentRow=sheet1.createRow(RowNum);
					for(int i=0;i<str.length;i++){
						currentRow.createCell(i).setCellValue(str[i]);
					}
				}
				FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress1);
				workBook.write(fileOutputStream);
				fileOutputStream.close();
				br.close();
				System.out.println("Done");
			} catch (Exception ex) {System.out.println(ex.getMessage()+"Exception in try");}
	try {
  	Workbook wba1 = new Workbook();
      wba1.loadFromFile(xlsxFileAddress1);
      Worksheet sheet1 = wba1.getWorksheets().get(0);        
      sheet1.getCellRange("G1").setValue("2018");
      sheet1.getCellRange("H1").setValue("2019");
      sheet1.getCellRange("I1").setValue("2020");
      sheet1.getCellRange("J1").setValue("2021");
      sheet1.getCellRange("K1").setValue("2022");
      sheet1.getCellRange("F2").setValue("January");
      sheet1.getCellRange("F3").setValue("February");
      sheet1.getCellRange("F4").setValue("March");
      sheet1.getCellRange("F5").setValue("April");
      sheet1.getCellRange("F6").setValue("May");
      sheet1.getCellRange("F7").setValue("June");
      sheet1.getCellRange("F8").setValue("July");
      sheet1.getCellRange("F9").setValue("August");
      sheet1.getCellRange("F10").setValue("September");
      sheet1.getCellRange("F11").setValue("October");
      sheet1.getCellRange("F12").setValue("November");
      sheet1.getCellRange("F13").setValue("December");
      sheet1.getCellRange("D2:D65").setFormula("=INDIRECT(\"A\"&ROW())&INDIRECT(\"B\"&ROW())");
      sheet1.getCellRange("E2:E65").setFormula("=SUBSTITUTE(INDIRECT(\"C\"&ROW()),CHAR(34),\"\")");
      sheet1.getCellRange("G2:G13").setFormula("=IFERROR(VLOOKUP(INDIRECT(\"F\"&ROW())&INDIRECT(\"G1\"),$D:$E,2,0),\"\")");
      sheet1.getCellRange("H2:H13").setFormula("=IFERROR(VLOOKUP(INDIRECT(\"F\"&ROW())&INDIRECT(\"H1\"),$D:$E,2,0),\"\")");
      sheet1.getCellRange("I2:I13").setFormula("=IFERROR(VLOOKUP(INDIRECT(\"F\"&ROW())&INDIRECT(\"I1\"),$D:$E,2,0),\"\")");
      sheet1.getCellRange("J2:J13").setFormula("=IFERROR(VLOOKUP(INDIRECT(\"F\"&ROW())&INDIRECT(\"J1\"),$D:$E,2,0),\"\")");
      sheet1.getCellRange("K2:K13").setFormula("=IFERROR(VLOOKUP(INDIRECT(\"F\"&ROW())&INDIRECT(\"K1\"),$D:$E,2,0),\"\")");
      wba1.getCalculationMode();
      wba1.calculateAllValue();
      wba1.calculateFormulaValue(null);
//      wba1.saveToFile(downloadFilepath + "//1.4check.xlsx", ExcelVersion.Version2013);
      String firstyear = "";
      String secondyear = "";
      String thriedyear = "";
      String forthyear = "";
      String fiveyear = "";
      String midvalue = "";
      for (int j = 2; j <= 14; j++) {
      	String aa = sheet1.getRange().get("G" + j).getDisplayedText();
      	String ab = sheet1.getRange().get("H" + j).getDisplayedText();
      	String ac = sheet1.getRange().get("I" + j).getDisplayedText();
      	String ad = sheet1.getRange().get("J" + j).getDisplayedText();
      	String ae = sheet1.getRange().get("K" + j).getDisplayedText();
          firstyear = firstyear + midvalue + aa;
          secondyear = secondyear + midvalue + ab;
          thriedyear = thriedyear + midvalue + ac;
          forthyear = forthyear + midvalue + ad;
          fiveyear = fiveyear + midvalue + ae; 
          if(midvalue.contentEquals("")) {
          	midvalue = "^";
          }
      }
      sheet.updateCellvalue(sheetsService, sheetname + "!A87", firstyear,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A88", secondyear,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A89", thriedyear,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A90", forthyear,sheetid);
      sheet.updateCellvalue(sheetsService, sheetname + "!A91", fiveyear,sheetid);        
  	} catch(Exception e1) {}  	 
		}
		

//4.3
				String csvFile2 = downloadFilepath +  "4.3.csv"; //csv file address
				String xlsxFile2 = downloadFilepath +   "4.3.xlsx"; //xlsx file address
				File fa111 = new File(csvFile2);
				for (int i1 = 2; i1 < 150; i1++) {
					Thread.sleep(1000);
					if(fa111.exists()){
						i1 = 150;
					}	
				}
				if(fa111.exists()){
					int RowNum=-1;
					try {
						@SuppressWarnings("resource")
						XSSFWorkbook workBook = new XSSFWorkbook();
						XSSFSheet sheet1 = workBook.createSheet("sheet1");
						String currentLine=null;
						RowNum=-1;
						BufferedReader br = new BufferedReader(new FileReader(csvFile2));
						while ((currentLine = br.readLine()) != null) {
							String str[] = currentLine.split((",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"));
							RowNum++;
							XSSFRow currentRow=sheet1.createRow(RowNum);
							for(int i=0;i<str.length;i++){
								currentRow.createCell(i).setCellValue(str[i]);
							}
						}				
						FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFile2);
						workBook.write(fileOutputStream);
						fileOutputStream.close();
						br.close();
						System.out.println("Done");
					} catch (Exception ex) {System.out.println(ex.getMessage()+"Exception in try");}
			try {
		    	Workbook wba1 = new Workbook();
		        wba1.loadFromFile(xlsxFile2);
		        Worksheet sheet2 = wba1.getWorksheets().get(0);
		        sheet2.getCellRange("C2:C2").setFormula("=SUBSTITUTE(A2&\"^\"&A3&\"^\"&A4,CHAR(34),\"\")");
		        sheet2.getCellRange("C3:C3").setFormula("=SUBSTITUTE(B2&\"^\"&B3&\"^\"&B4,CHAR(34),\"\")");
		        wba1.getCalculationMode();
		        wba1.calculateAllValue();
		        wba1.calculateFormulaValue(null);
		        String val1 = sheet2.getRange().get("C2").getDisplayedText();
		        String val2 = sheet2.getRange().get("C3").getDisplayedText();
		        sheet.updateCellvalue(sheetsService, sheetname + "!H94", val1,sheetid);
		        sheet.updateCellvalue(sheetsService, sheetname + "!I94", val2,sheetid);
		    	} catch(Exception e1) {}  	 
				}}}
  }
  @BeforeTest
  public void beforeTest() {
  }

  @AfterTest
  public void afterTest() {
  }

}
