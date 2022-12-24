package com.sn.ajr.citation; 
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.HashMap;
import java.util.List;
//import javax.xml.ws.Response;
import org.openqa.selenium.chrome.ChromeOptions;
import com.google.api.services.sheets.v4.Sheets;
import com.sn.ajr.ProjGSheet;
import com.spire.xls.ExcelVersion;
import com.spire.xls.OrderBy;
import com.spire.xls.SortComparsionType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.sorting.SortColumn;
import jakarta.servlet.http.HttpServletRequest;
public class Ccalculation {
//	spire.xls
	public static void classname() throws InterruptedException, IOException, GeneralSecurityException { 
		System.out.println("WOS calculation start");
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
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		ProjGSheet sheet = new ProjGSheet();
		 Sheets sheetsService = sheet.getSheetsService("Google Sheets Example");
		 
//web of science
try{
String catsheet = downloadFilepath + "citaion1.xlsx";
String catsheet1 = downloadFilepath + "citaion2.xlsx";
String catsheet2 = downloadFilepath + "citaion3.xlsx";
String catsheet3 = downloadFilepath + "citaion4.xlsx";
String catsheet4 = downloadFilepath + "citaion5.xlsx";

String typesheet = downloadFilepath + "savedrecs.xlsx";
String typesheet1 = downloadFilepath + "savedrecs (1).xlsx";
String typesheet2 = downloadFilepath + "savedrecs (2).xlsx";
String typesheet3 = downloadFilepath + "savedrecs (3).xlsx";
String typesheet4 = downloadFilepath + "savedrecs (4).xlsx";

Workbook wbmaster = new Workbook();
wbmaster.loadFromFile(catsheet);
Worksheet catsheeta = wbmaster.getWorksheets().get(0);

try {
Workbook wbb1 = new Workbook();
wbb1.loadFromFile(catsheet1);
Worksheet sheetb1 = wbb1.getWorksheets().get(0); 
for (int i1 = 12; i1 < 1012; i1++) {
 int i2 = i1 + 1000;
 String valuess = sheetb1.getRange().get("A" + i1).getDisplayedText();
 if (!valuess.isEmpty()) {
catsheeta.getCellRange("A" + i2).setValue(sheetb1.getRange().get("A" + i1).getDisplayedText());}
String valuess1 = sheetb1.getRange().get("B" + i1).getDisplayedText();
if (!valuess1.isEmpty()) {
catsheeta.getCellRange("B" + i2).setValue(sheetb1.getRange().get("B" + i1).getDisplayedText());}
String valuess2 = sheetb1.getRange().get("H" + i1).getDisplayedText();
if (!valuess2.isEmpty()) {
catsheeta.getCellRange("H" + i2).setValue(sheetb1.getRange().get("H" + i1).getDisplayedText());}
String valuess3 = sheetb1.getRange().get("Q" + i1).getDisplayedText();
if (!valuess3.isEmpty()) {
catsheeta.getCellRange("Q" + i2).setValue(sheetb1.getRange().get("Q" + i1).getDisplayedText());}
String valuess4 = sheetb1.getRange().get("T" + i1).getDisplayedText();
if (!valuess4.isEmpty()) {
catsheeta.getCellRange("T" + i2).setValue(sheetb1.getRange().get("T" + i1).getDisplayedText());}
String valuess5 = sheetb1.getRange().get("BK" + i1).getDisplayedText();
if (!valuess5.isEmpty()) {
catsheeta.getCellRange("BK" + i2).setValue(sheetb1.getRange().get("BK" + i1).getDisplayedText());}
}
} catch(Exception e1) {}
try {
Workbook wbb1 = new Workbook();
wbb1.loadFromFile(catsheet2);
Worksheet sheetb1 = wbb1.getWorksheets().get(0); 
 for (int i1 = 12; i1 < 1012; i1++) {
	 int i2 = i1 + 2000;
	 String valuess = sheetb1.getRange().get("A" + i1).getDisplayedText();
	 if (!valuess.isEmpty()) {
catsheeta.getCellRange("A" + i2).setValue(sheetb1.getRange().get("A" + i1).getDisplayedText());}
 String valuess1 = sheetb1.getRange().get("B" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("B" + i2).setValue(sheetb1.getRange().get("B" + i1).getDisplayedText());}
 String valuess2 = sheetb1.getRange().get("H" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("H" + i2).setValue(sheetb1.getRange().get("H" + i1).getDisplayedText());}
 String valuess3 = sheetb1.getRange().get("Q" + i1).getDisplayedText();
 if (!valuess3.isEmpty()) {
catsheeta.getCellRange("Q" + i2).setValue(sheetb1.getRange().get("Q" + i1).getDisplayedText());}
 String valuess4 = sheetb1.getRange().get("T" + i1).getDisplayedText();
 if (!valuess4.isEmpty()) {
catsheeta.getCellRange("T" + i2).setValue(sheetb1.getRange().get("T" + i1).getDisplayedText());}
 String valuess5 = sheetb1.getRange().get("BK" + i1).getDisplayedText();
 if (!valuess5.isEmpty()) {
catsheeta.getCellRange("BK" + i2).setValue(sheetb1.getRange().get("BK" + i1).getDisplayedText());}
 }
 } catch(Exception e1) {}
try {
Workbook wbb1 = new Workbook();
wbb1.loadFromFile(catsheet3);
Worksheet sheetb1 = wbb1.getWorksheets().get(0); 
 for (int i1 = 12; i1 < 1012; i1++) {
	 int i2 = i1 + 3000;
	 String valuess = sheetb1.getRange().get("A" + i1).getDisplayedText();
	 if (!valuess.isEmpty()) {
catsheeta.getCellRange("A" + i2).setValue(sheetb1.getRange().get("A" + i1).getDisplayedText());}
 String valuess1 = sheetb1.getRange().get("B" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("B" + i2).setValue(sheetb1.getRange().get("B" + i1).getDisplayedText());}
 String valuess2 = sheetb1.getRange().get("H" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("H" + i2).setValue(sheetb1.getRange().get("H" + i1).getDisplayedText());}
 String valuess3 = sheetb1.getRange().get("Q" + i1).getDisplayedText();
 if (!valuess3.isEmpty()) {
catsheeta.getCellRange("Q" + i2).setValue(sheetb1.getRange().get("Q" + i1).getDisplayedText());}
 String valuess4 = sheetb1.getRange().get("T" + i1).getDisplayedText();
 if (!valuess4.isEmpty()) {
catsheeta.getCellRange("T" + i2).setValue(sheetb1.getRange().get("T" + i1).getDisplayedText());}
 String valuess5 = sheetb1.getRange().get("BK" + i1).getDisplayedText();
 if (!valuess5.isEmpty()) {
catsheeta.getCellRange("BK" + i2).setValue(sheetb1.getRange().get("BK" + i1).getDisplayedText());}
 }
 } catch(Exception e1) {}
try {
	
Workbook wbb1 = new Workbook();
wbb1.loadFromFile(catsheet4);
Worksheet sheetb1 = wbb1.getWorksheets().get(0); 
 for (int i1 = 12; i1 < 1012; i1++) {
	 int i2 = i1 + 4000;
	 String valuess = sheetb1.getRange().get("A" + i1).getDisplayedText();
	 if (!valuess.isEmpty()) {
catsheeta.getCellRange("A" + i2).setValue(sheetb1.getRange().get("A" + i1).getDisplayedText());}
 String valuess1 = sheetb1.getRange().get("B" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("B" + i2).setValue(sheetb1.getRange().get("B" + i1).getDisplayedText());}
 String valuess2 = sheetb1.getRange().get("H" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("H" + i2).setValue(sheetb1.getRange().get("H" + i1).getDisplayedText());}
 String valuess3 = sheetb1.getRange().get("Q" + i1).getDisplayedText();
 if (!valuess3.isEmpty()) {
catsheeta.getCellRange("Q" + i2).setValue(sheetb1.getRange().get("Q" + i1).getDisplayedText());}
 String valuess4 = sheetb1.getRange().get("T" + i1).getDisplayedText();
 if (!valuess4.isEmpty()) {
catsheeta.getCellRange("T" + i2).setValue(sheetb1.getRange().get("T" + i1).getDisplayedText());}
 String valuess5 = sheetb1.getRange().get("BK" + i1).getDisplayedText();
 if (!valuess5.isEmpty()) {
catsheeta.getCellRange("BK" + i2).setValue(sheetb1.getRange().get("BK" + i1).getDisplayedText());}
 }
 } catch(Exception e1) {}

try {
Workbook wba1 = new Workbook();
wba1.loadFromFile(typesheet);
Worksheet sheeta1 = wba1.getWorksheets().get(0); 
for (int i1 = 1; i1 < 1002; i1++) {
 int i2 = i1 + 10;
 String valuess1 = sheeta1.getRange().get("I" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("BM" + i2).setValue(sheeta1.getRange().get("I" + i1).getDisplayedText());}
 String valuess2 = sheeta1.getRange().get("N" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("BN" + i2).setValue(sheeta1.getRange().get("N" + i1).getDisplayedText());}
}
} catch(Exception e1) {}

try {
Workbook wba1 = new Workbook();
wba1.loadFromFile(typesheet1);
Worksheet sheeta1 = wba1.getWorksheets().get(0); 
for (int i1 = 2; i1 < 2002; i1++) {
 int i2 = i1 + 1010;
 String valuess1 = sheeta1.getRange().get("I" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("BM" + i2).setValue(sheeta1.getRange().get("I" + i1).getDisplayedText());}
 String valuess2 = sheeta1.getRange().get("N" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("BN" + i2).setValue(sheeta1.getRange().get("N" + i1).getDisplayedText());}
}
} catch(Exception e1) {}
try {
Workbook wba1 = new Workbook();
wba1.loadFromFile(typesheet2);
Worksheet sheeta1 = wba1.getWorksheets().get(0); 
for (int i1 = 2; i1 < 3002; i1++) {
 int i2 = i1 + 2010;
 String valuess1 = sheeta1.getRange().get("I" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("BM" + i2).setValue(sheeta1.getRange().get("I" + i1).getDisplayedText());}
 String valuess2 = sheeta1.getRange().get("N" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("BN" + i2).setValue(sheeta1.getRange().get("N" + i1).getDisplayedText());}
}
} catch(Exception e1) {} 
try {
Workbook wba1 = new Workbook();
wba1.loadFromFile(typesheet3);
Worksheet sheeta1 = wba1.getWorksheets().get(0); 
for (int i1 = 2; i1 < 4002; i1++) {
 int i2 = i1 + 3010;
 String valuess1 = sheeta1.getRange().get("I" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("BM" + i2).setValue(sheeta1.getRange().get("I" + i1).getDisplayedText());}
 String valuess2 = sheeta1.getRange().get("N" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("BN" + i2).setValue(sheeta1.getRange().get("N" + i1).getDisplayedText());}
}
} catch(Exception e1) {} 
try {
Workbook wba1 = new Workbook();
wba1.loadFromFile(typesheet4);
Worksheet sheeta1 = wba1.getWorksheets().get(0); 
for (int i1 = 2; i1 < 5002; i1++) {
 int i2 = i1 + 4010;
 String valuess1 = sheeta1.getRange().get("I" + i1).getDisplayedText();
 if (!valuess1.isEmpty()) {
catsheeta.getCellRange("BM" + i2).setValue(sheeta1.getRange().get("I" + i1).getDisplayedText());}
 String valuess2 = sheeta1.getRange().get("N" + i1).getDisplayedText();
 if (!valuess2.isEmpty()) {
catsheeta.getCellRange("BN" + i2).setValue(sheeta1.getRange().get("N" + i1).getDisplayedText());}
}
} catch(Exception e1) {} 

int tstring2 = 1;
for (int i1 = 12; i1 < 100000; i1++) {
	if (!catsheeta.getRange().get("A" + i1).getDisplayedText().contentEquals("")) 
	{tstring2 = tstring2 + 1;}else {
		i1 = 100000;
	}
}
tstring2 = tstring2 + 10;
catsheeta.getCellRange("BO12:BO" + tstring2).setFormula("IF(INDIRECT(\"A\"&ROW())=\"\",\"\",VLOOKUP(INDIRECT(\"A\"&ROW()),BM:BN,2,0))");
catsheeta.getCellRange("BP12:BP" + tstring2).setFormula("IF(INDIRECT(\"BK\"&ROW())=\"\",\"\",IF(INDIRECT(\"BK\"&ROW())=\"no\",0,COUNTIF(INDIRECT(\"BK11:BK\"&ROW()),INDIRECT(\"BK\"&ROW()))))");
catsheeta.getCellRange("BQ12:BQ" + tstring2).setFormula("IF(INDIRECT(\"BP\"&ROW())=1,COUNTIF(INDIRECT(\"BK:BK\"),INDIRECT(\"BK\"&ROW())))"); 
SortColumn column = wbmaster.getDataSorter().getSortColumns().add(62, SortComparsionType.Values, OrderBy.Descending);
wbmaster.getDataSorter().sort(catsheeta.getCellRange("A11:BN" + tstring2));

catsheeta.getCellRange("BA1").setValue("*Article*");
catsheeta.getCellRange("BA2").setValue("*Review*");
catsheeta.getCellRange("BA3").setValue("*Proceedings Paper*");
catsheeta.getCellRange("BB1").setFormula("IFERROR(COUNTIFS(INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")-INDIRECT(\"BB3\")");
catsheeta.getCellRange("BB2").setFormula("IFERROR(COUNTIFS(INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")");
catsheeta.getCellRange("BB3").setFormula("IFERROR(COUNTIFS(INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")");
catsheeta.getCellRange("BC1").setFormula("IFERROR(SUMIFS(INDIRECT(\"BK:BK\"),INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")-INDIRECT(\"BC3\")");
catsheeta.getCellRange("BC2").setFormula("IFERROR(SUMIFS(INDIRECT(\"BK:BK\"),INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")");
catsheeta.getCellRange("BC3").setFormula("IFERROR(SUMIFS(INDIRECT(\"BK:BK\"),INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")");
catsheeta.getCellRange("BC4").setFormula("SUM(BC1:BC3)");
catsheeta.getCellRange("BD1:BD3").setFormula("IFERROR(AVERAGEIFS(INDIRECT(\"BK12:BK5100\"), INDIRECT(\"H12:H5100\"), \">=2019\", INDIRECT(\"H12:H5100\"), \"<=2022\",INDIRECT(\"BO12:BO5100\"),INDIRECT(\"BA\"&ROW())),\"\")");
catsheeta.getCellRange("BE1:BE3").setFormula("IFERROR(IF(INDIRECT(\"BC\"&ROW())=0,0,INDIRECT(\"BC\"&ROW())/INDIRECT(\"BC4\")),0)");
catsheeta.getCellRange("BF1").setFormula("IFERROR(COUNTIFS(INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\",INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"BK:BK\"),\"=0\"),\"\")-INDIRECT(\"BF3\")");
catsheeta.getCellRange("BF2").setFormula("IFERROR(COUNTIFS(INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\",INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"BK:BK\"),\"=0\"),\"\")");
catsheeta.getCellRange("BF3").setFormula("IFERROR(COUNTIFS(INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\",INDIRECT(\"BO:BO\"),INDIRECT(\"BA\"&ROW()),INDIRECT(\"BK:BK\"),\"=0\"),\"\")");
catsheeta.getCellRange("BG1:BG3").setFormula("IFERROR(INDIRECT(\"BF\"&ROW())/INDIRECT(\"BB\"&ROW()),\"\")");
catsheeta.getCellRange("BH3").setFormula("IFERROR(COUNTIFS(INDIRECT(\"H:H\"),\">=2019\",INDIRECT(\"H:H\"),\"<=2022\"),\"\")");

catsheeta.getCellRange("BR2").setValue("2016");
catsheeta.getCellRange("BR3").setValue("2017");
catsheeta.getCellRange("BR4").setValue("2018");
catsheeta.getCellRange("BR5").setValue("2019");
catsheeta.getCellRange("BR6").setValue("2020");
catsheeta.getCellRange("BS2:BS5").setFormula("COUNTIFS(INDIRECT(\"H:H\"),INDIRECT(\"BR\"&ROW()))");
catsheeta.getCellRange("BS6").setFormula("COUNTIFS(INDIRECT(\"H:H\"),INDIRECT(\"BR\"&ROW()))+COUNTIFS(INDIRECT(\"H:H\"),\"2021\")+COUNTIFS(INDIRECT(\"H:H\"),\"2022\")");
catsheeta.getCellRange("BT2:BT6").setFormula("COUNTIFS(INDIRECT(\"H:H\"),INDIRECT(\"BR\"&ROW()),INDIRECT(\"BK:BK\"),\"0\")");
 
wbmaster.getCalculationMode();
wbmaster.calculateAllValue();
wbmaster.calculateFormulaValue(null);
String medvalue = "^";
String fiveyearsited = "";
String fiveyearzerosited = "";
for (int i12 = 2; i12 < 7; i12++) {
	String res1a = catsheeta.getRange().get("BS" + i12).getDisplayedText();
	String res1b = catsheeta.getRange().get("BT" + i12).getDisplayedText();
	if (fiveyearsited.equals("")) {
		medvalue = "";
		}else {
		medvalue= "^";
		}	
	fiveyearsited = fiveyearsited + medvalue + res1b;
	fiveyearzerosited = fiveyearzerosited + medvalue + res1a;
}

String artiles = "";
String reviews = "";
String rpoceedingpapers = ""; 
for (int i12 = 54; i12 < 60; i12++) {
String resc1 = catsheeta.getRange().get(1, i12).getDisplayedText();
String resc2 = catsheeta.getRange().get(2, i12).getDisplayedText();
String resc3 = catsheeta.getRange().get(3, i12).getDisplayedText();
if (artiles.equals("")) {
medvalue = "";
}else {
medvalue= "^";
}
artiles = artiles + medvalue + resc1;
reviews = reviews + medvalue + resc2;
rpoceedingpapers = rpoceedingpapers + medvalue + resc3;
}
String twoyearvalue = catsheeta.getRange().get("BH3").getDisplayedText();
for (int i13 = 12; i13 < 5012; i13++) {
	
	String yearcheck = catsheeta.getCellRange("H" + i13).getDisplayedText();
	if(yearcheck.contentEquals("2016") ) {
		catsheeta.getCellRange("BK" + i13).setValue("no");
	}
	if(yearcheck.contentEquals("2017") ) {
		catsheeta.getCellRange("BK" + i13).setValue("no");
	}
	if(yearcheck.contentEquals("2018") ) {
		catsheeta.getCellRange("BK" + i13).setValue("no");
	}
	if(yearcheck.contentEquals("2021") ) {
		catsheeta.getCellRange("BK" + i13).setValue("no");
	}
	if(yearcheck.contentEquals("2022") ) {
		catsheeta.getCellRange("BK" + i13).setValue("no");
	}
}
String col1 = "";
String col2 = "";
String col3 = "";
String col4 = "";
String col5 = "";
String col6 = "";
String col7 = "";
int i14 = 12;
for (int i13 = 12; i14 < 24;) {
String checkvalue1 = catsheeta.getRange().get("BK" + i13).getDisplayedText();	
if (!checkvalue1.contentEquals("no")) {	
String res1a = catsheeta.getRange().get("A" + i13).getDisplayedText();
String res1b = catsheeta.getRange().get("B" + i13).getDisplayedText();
String res1c = catsheeta.getRange().get("BO" + i13).getDisplayedText();
String res1d = catsheeta.getRange().get("H" + i13).getDisplayedText();
String res1e = catsheeta.getRange().get("Q" + i13).getDisplayedText();
String res1f = catsheeta.getRange().get("T" + i13).getDisplayedText();
String res1g = catsheeta.getRange().get("BK" + i13).getDisplayedText();
if (col1.equals("")) {
medvalue = "";
}else {
medvalue= "^";
}
col1 = col1 + medvalue + res1a;
col2 = col2 + medvalue + res1b;
col3 = col3 + medvalue + res1c;
col4 = col4 + medvalue + res1d;
col5 = col5 + medvalue + res1e;
col6 = col6 + medvalue + res1f;
col7 = col7 + medvalue + res1g;
i13++;
i14++;
}else {i13++;}
}

wbmaster.getCalculationMode();
wbmaster.calculateAllValue();
wbmaster.calculateFormulaValue(null);
String Frequencyone = "";
String Frequencytwo = ""; 
for (int i13 = 12; i13 < 5012; i13++) {
	
	String checkvalue = catsheeta.getRange().get("BP" + i13).getDisplayedText(); 
	if (checkvalue.contentEquals("1")) {
		String resc1 = catsheeta.getRange().get("BK" + i13).getDisplayedText();
		String resc2 = catsheeta.getRange().get("BQ" + i13).getDisplayedText();
		if (Frequencyone.equals("")) {
			medvalue = "";
			}else {
			medvalue= "^";
			}
		Frequencyone = Frequencyone + medvalue + resc1;
		Frequencytwo = Frequencytwo + medvalue + resc2;
	}
}
//System.out.println("ACitationType|" + artiles + "\n" + reviews + "\n" + rpoceedingpapers + "|" + col1 + "|" + col2 + "|" + col3 + "|" + col4 + "|" + col5 + "|" + col6 + "|" + col7);
//System.out.println("AFrequencyofarticlesA|" + twoyearvalue + "|" + fiveyearzerosited + "|" + fiveyearsited + "|" + Frequencyone + "|" + Frequencytwo);
Thread.sleep(500);
sheet.updateCellvalue(sheetsService, sheetname + "!H12", "ACitationType|" + artiles + "\n" + reviews + "\n" + rpoceedingpapers + "|" + col1 + "|" + col2 + "|" + col3 + "|" + col4 + "|" + col5 + "|" + col6 + "|" + col7,sheetid);
sheet.updateCellvalue(sheetsService, sheetname + "!E165", "AFrequencyofarticlesA|" + twoyearvalue + "|" + fiveyearzerosited + "|" + fiveyearsited + "|" + Frequencyone + "|" + Frequencytwo,sheetid);
wbmaster.saveToFile("C:\\AJR Self-service tool" + "\\Citationdata.xlsx", ExcelVersion.Version2013);
System.out.println("WOS Done");
} catch(Exception e1) {System.out.println("excelptioninca");} }
	}}
}