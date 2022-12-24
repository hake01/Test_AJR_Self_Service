package com.sn.ajr.jcr;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.slides.v1.Slides;
import com.sn.ajr.ProjGSheet;
import com.sn.ajr.ProjGSlide;
import com.sn.ajr.UpdateSlidePojo;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
public class ProjJcrCategory4 {
    private static final String APPLICATION_NAME = "Google Slides API Java Quickstart";
		public static String mainclass() throws IOException, GeneralSecurityException {
			System.out.println("cat4 start");
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
//			String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
			ProjGSheet sheetq = new ProjGSheet();
			Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example");    
		  	String sheetname = "";
		  	String pptid = "";
		  	try {
		  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
				for (List<?> row : values) {
			sheetname = (String) row.get(0); 
			pptid = (String) row.get(11); 
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
			ProjGSlide pSlide = new ProjGSlide();
			Slides service = pSlide.getSlideService(APPLICATION_NAME);

//		File currentDirFile = new File(".");
//		String rootpath1 = currentDirFile.getAbsolutePath();
//		String rootpath = rootpath1.replace(".", "");
//		String downloadFilepath = rootpath + "DownloadingTempfiles\\";
		
    	try {
    	String xlsxFilecat1 = downloadFilepath + "JCR_category4.xlsx"; //xlsx file address
    	File fa1 = new File(xlsxFilecat1);
    	if(fa1.exists()){
    		String rank = "1";
    		try{
    			List<List<Object>> values = sheet.getCellData(sheetsService, sheetname + "!S82",sheetid);
    			for (List<?> row : values) {
    				rank = (String) row.get(0);
    			}		
    	} catch(Exception e1) {}
    	Workbook wbc1 = new Workbook();
        wbc1.loadFromFile(xlsxFilecat1);
        Worksheet sheetd1 = wbc1.getWorksheets().get(0);        		
    		String cate1 = "";
           for (int i1 = 4; i1 < 24; i1++) {
           	String value = sheetd1.getRange().get("C" + i1).getValue();
           	if (value.contentEquals("N/A")) {
                 sheetd1.getCellRange("C" + i1).setValue(sheetd1.getCellRange("D" + i1).getValue());
           	}
                 String re1 = sheetd1.getRange().get("C" + i1).getDisplayedText();
                 if (cate1.contentEquals("")) {
                	 cate1 = re1;
                 }else {
                 cate1 = cate1 + "^" + re1;}
           	
                }
           
           int ranka=Integer.parseInt(rank);
           if (ranka > 20) {
           String value = sheetd1.getRange().get("C" + ranka).getValue();
        	if (value.contentEquals("N/A")) {
              sheetd1.getCellRange("C" + ranka).setValue(sheetd1.getCellRange("D" + ranka).getValue());
        	}     String re1 = sheetd1.getRange().get("C" + ranka).getDisplayedText();
              cate1 = cate1 + "^" + re1;
        }
     sheet.updateCellvalue(sheetsService, sheetname + "!T80", cate1,sheetid);
     
           sheetd1.getRange().get("K9:L1000").setNumberFormat("0.000");
           sheetd1.getRange().get("G9:G1000").setNumberFormat("0.000");
           sheetd1.getRange().get("F9:F1000").setNumberFormat("0,000");
          int aa = 34;
          int ab = 4; 
          List<UpdateSlidePojo> updateSlides = new ArrayList<>(); 
          UpdateSlidePojo rankvalue1a = new UpdateSlidePojo(65,"<u32>", rank); 
          	ranka = ranka + 3;
        	UpdateSlidePojo call1 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call2 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call3 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call4 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call5 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call6 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call7 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call8 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call9 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call10 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call11 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call12 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call13 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call14 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call15 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call16 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call17 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call18 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call19 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call20 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call21 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("B" + ranka).getDisplayedText());aa++;ab++;

        	aa = 78;
            ab = 4;
    		UpdateSlidePojo calld1 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld2 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld3 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld4 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld5 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld6 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld7 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld8 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld9 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld10 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld11 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld12 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld13 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld14 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld15 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld16 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld17 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld18 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld19 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld20 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo calld21 = new UpdateSlidePojo(65,"<u" + aa + ">",sheetd1.getRange().get("C" + ranka).getDisplayedText());aa++;ab++;
        	aa = 99;
            ab = 4;
    		UpdateSlidePojo callf1 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf2 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf3 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf4 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf5 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf6 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf7 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf8 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf9 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf10 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf11 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf12 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf13 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf14 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf15 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf16 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf17 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf18 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf19 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf20 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo callf21 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("F" + ranka).getDisplayedText());aa++;ab++;        	
        	
        aa = 120;
        ab = 4;  
		UpdateSlidePojo callg1 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg2 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg3 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg4 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg5 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg6 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg7 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg8 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg9 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg10 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg11 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg12 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg13 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg14 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg15 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg16 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg17 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg18 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg19 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg20 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callg21 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("G" + ranka).getDisplayedText());aa++;ab++;
      	aa = 141;
          ab = 4;
  		UpdateSlidePojo calll1 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll2 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll3 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll4 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll5 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll6 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll7 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll8 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll9 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll10 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll11 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll12 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll13 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll14 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll15 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll16 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll17 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll18 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll19 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll20 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo calll21 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("L" + ranka).getDisplayedText());aa++;ab++;
      	aa = 162;
          ab = 4;
  		UpdateSlidePojo callk1 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk2 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk3 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk4 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk5 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk6 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk7 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk8 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk9 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk10 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk11 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk12 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk13 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk14 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk15 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk16 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk17 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk18 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk19 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk20 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ab).getDisplayedText());aa++;ab++;
    	UpdateSlidePojo callk21 = new UpdateSlidePojo(65,"<u0" + aa + ">",sheetd1.getRange().get("K" + ranka).getDisplayedText());aa++;ab++;      
      	
    	updateSlides.add(rankvalue1a); 
    	updateSlides.add(call1);
      	updateSlides.add(call2);
      	updateSlides.add(call3);
      	updateSlides.add(call4); 
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
      	updateSlides.add(calld1);
      	updateSlides.add(calld2);
      	updateSlides.add(calld3);
      	updateSlides.add(calld4); 
      	updateSlides.add(calld4); 
      	updateSlides.add(calld5); 
      	updateSlides.add(calld6); 
      	updateSlides.add(calld7); 
      	updateSlides.add(calld8); 
      	updateSlides.add(calld9); 
      	updateSlides.add(calld10); 
      	updateSlides.add(calld11); 
      	updateSlides.add(calld12); 
      	updateSlides.add(calld13); 
      	updateSlides.add(calld14); 
      	updateSlides.add(calld15); 
      	updateSlides.add(calld16); 
      	updateSlides.add(calld17); 
      	updateSlides.add(calld18); 
      	updateSlides.add(calld19); 
      	updateSlides.add(calld20); 
      	updateSlides.add(calld21); 
      	updateSlides.add(callf1);
      	updateSlides.add(callf2);
      	updateSlides.add(callf3);
      	updateSlides.add(callf4); 
      	updateSlides.add(callf4);
      	updateSlides.add(callf5); 
      	updateSlides.add(callf6); 
      	updateSlides.add(callf7); 
      	updateSlides.add(callf8); 
      	updateSlides.add(callf9); 
      	updateSlides.add(callf10); 
      	updateSlides.add(callf11); 
      	updateSlides.add(callf12); 
      	updateSlides.add(callf13); 
      	updateSlides.add(callf14); 
      	updateSlides.add(callf15); 
      	updateSlides.add(callf16); 
      	updateSlides.add(callf17); 
      	updateSlides.add(callf18); 
      	updateSlides.add(callf19); 
      	updateSlides.add(callf20); 
      	updateSlides.add(callf21); 
      	updateSlides.add(callg1);
      	updateSlides.add(callg2);
      	updateSlides.add(callg3);
      	updateSlides.add(callg4); 
      	updateSlides.add(callg4);  
      	updateSlides.add(callg5); 
      	updateSlides.add(callg6); 
      	updateSlides.add(callg7); 
      	updateSlides.add(callg8); 
      	updateSlides.add(callg9); 
      	updateSlides.add(callg10); 
      	updateSlides.add(callg11); 
      	updateSlides.add(callg12); 
      	updateSlides.add(callg13); 
      	updateSlides.add(callg14); 
      	updateSlides.add(callg15); 
      	updateSlides.add(callg16); 
      	updateSlides.add(callg17); 
      	updateSlides.add(callg18); 
      	updateSlides.add(callg19); 
      	updateSlides.add(callg20); 
      	updateSlides.add(callg21); 
      	updateSlides.add(calll1);
      	updateSlides.add(calll2);
      	updateSlides.add(calll3);
      	updateSlides.add(calll4); 
      	updateSlides.add(calll4);  
      	updateSlides.add(calll5); 
      	updateSlides.add(calll6); 
      	updateSlides.add(calll7); 
      	updateSlides.add(calll8); 
      	updateSlides.add(calll9); 
      	updateSlides.add(calll10); 
      	updateSlides.add(calll11); 
      	updateSlides.add(calll12); 
      	updateSlides.add(calll13); 
      	updateSlides.add(calll14); 
      	updateSlides.add(calll15); 
      	updateSlides.add(calll16); 
      	updateSlides.add(calll17); 
      	updateSlides.add(calll18); 
      	updateSlides.add(calll19); 
      	updateSlides.add(calll20); 
      	updateSlides.add(calll21); 
      	updateSlides.add(callk1);
      	updateSlides.add(callk2);
      	updateSlides.add(callk3);
      	updateSlides.add(callk4); 
      	updateSlides.add(callk4);  
      	updateSlides.add(callk5); 
      	updateSlides.add(callk6); 
      	updateSlides.add(callk7); 
      	updateSlides.add(callk8); 
      	updateSlides.add(callk9); 
      	updateSlides.add(callk10); 
      	updateSlides.add(callk11); 
      	updateSlides.add(callk12); 
      	updateSlides.add(callk13); 
      	updateSlides.add(callk14); 
      	updateSlides.add(callk15); 
      	updateSlides.add(callk16); 
      	updateSlides.add(callk17); 
      	updateSlides.add(callk18); 
      	updateSlides.add(callk19); 
      	updateSlides.add(callk20); 
      	updateSlides.add(callk21); 
      	pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides);
		}
      	        } catch(Exception e1) {}}
    	System.out.println("cat4 end");}
    	return null;
    }

}