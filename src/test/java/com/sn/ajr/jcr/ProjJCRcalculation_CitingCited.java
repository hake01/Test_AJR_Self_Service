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

import jakarta.servlet.http.HttpServletRequest;
public class ProjJCRcalculation_CitingCited {
    private static final String APPLICATION_NAME = "Google Slides API Java Quickstart";
		public static String mainclass() throws IOException, GeneralSecurityException {
		System.out.println("cited citing start");
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
//		String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example");  
	  	String pptid = "";
	  	try {
	  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:P2",sheetidq); 
			for (List<?> row : values) { 
		pptid = (String) row.get(11); 
			}
	  	} catch (Exception e1) {}
	  	if(!pptid.contentEquals("")) {

	  		ProjGSlide pSlide = new ProjGSlide();
    	Slides service = pSlide.getSlideService(APPLICATION_NAME);
    	try {
            try {
            	File directoryq = new File(downloadFilepath);
                File[] filesq = directoryq.listFiles();
                for (File fq : filesq)
                	
                {
                	if (fq.getName().endsWith("Cited-Journal-Data-2021.xlsx")){
                		String oldfilename = fq.getName(); 
                    	File myfile = new File(downloadFilepath + oldfilename);
                    	myfile.renameTo(new File(downloadFilepath + "Cited.xlsx"));
                    	
                    }
                }
            } catch (Exception ex) {}   		

            String xlsxFilecitig = downloadFilepath + "Cited.xlsx";
            Workbook wbd1 = new Workbook();
            wbd1.loadFromFile(xlsxFilecitig);
           Worksheet sheetd1 = wbd1.getWorksheets().get(0);
         
           for (int i1 = 9; i1 < 30; i1++) {
           	String value = sheetd1.getRange().get("B" + i1).getValue();
                 sheetd1.getCellRange("B" + i1).setValue(value);
                if (!value.contentEquals("N/A")) {
                }else {
               	 sheetd1.getCellRange("B" + i1).setValue("");
                }
                }
           sheetd1.getRange().get("B9:B30").setNumberFormat("0.000");
          int aa = 13;
          int ab = 9; 
          List<UpdateSlidePojo> updateSlides = new ArrayList<>(); 
        	UpdateSlidePojo call1 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call2 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call3 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call4 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call5 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call6 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call7 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call8 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call9 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call10 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call11 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call12 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call13 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call14 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call15 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call16 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call17 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call18 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call19 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call20 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
        	aa = 33;
            ab = 7;
        	UpdateSlidePojo call21 = new UpdateSlidePojo(68,"<tcj" + aa + ">","ALL Journals");aa++;ab++;
        	UpdateSlidePojo call22 = new UpdateSlidePojo(68,"<tcj" + aa + ">","ALL OTHERS");aa++;ab++;
        	UpdateSlidePojo call23 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call24 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call25 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call26 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call27 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call28 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call29 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call30 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call31 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call32 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call33 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call34 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call35 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call36 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call37 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call38 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call39 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call40 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call41 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call42 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
        	aa = 55;
            ab = 7;
        	UpdateSlidePojo call43 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call44 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call45 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call46 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call47 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call48 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call49 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call50 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call51 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call52 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call53 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call54 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call55 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call56 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call57 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call58 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call59 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call60 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call61 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call62 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call63 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call64 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
        	aa = 77;
            ab = 7;
        	UpdateSlidePojo call65 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call66 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call67 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call68 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call69 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call70 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call71 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call72 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call73 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call74 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call75 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call76 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call77 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call78 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call79 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call80 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call81 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call82 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call83 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call84 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call85 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
        	UpdateSlidePojo call86 = new UpdateSlidePojo(68,"<tcj" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
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
        	updateSlides.add(call22); 
        	updateSlides.add(call23); 
        	updateSlides.add(call24); 
        	updateSlides.add(call25); 
        	updateSlides.add(call26); 
        	updateSlides.add(call27); 
        	updateSlides.add(call28); 
        	updateSlides.add(call29); 
        	updateSlides.add(call30); 
        	updateSlides.add(call31); 
        	updateSlides.add(call32); 
        	updateSlides.add(call33); 
        	updateSlides.add(call34); 
        	updateSlides.add(call35); 
        	updateSlides.add(call36); 
        	updateSlides.add(call37); 
        	updateSlides.add(call38); 
        	updateSlides.add(call39); 
        	updateSlides.add(call40);     
        	updateSlides.add(call41); 
        	updateSlides.add(call42); 
        	updateSlides.add(call43); 
        	updateSlides.add(call44); 
        	updateSlides.add(call45); 
        	updateSlides.add(call46); 
        	updateSlides.add(call47); 
        	updateSlides.add(call48); 
        	updateSlides.add(call49); 
        	updateSlides.add(call50); 
        	updateSlides.add(call51); 
        	updateSlides.add(call52); 
        	updateSlides.add(call53); 
        	updateSlides.add(call54); 
        	updateSlides.add(call55); 
        	updateSlides.add(call56); 
        	updateSlides.add(call57); 
        	updateSlides.add(call58); 
        	updateSlides.add(call59); 
        	updateSlides.add(call60); 
        	updateSlides.add(call61); 
        	updateSlides.add(call62); 
        	updateSlides.add(call63); 
        	updateSlides.add(call64); 
        	updateSlides.add(call65); 
        	updateSlides.add(call66); 
        	updateSlides.add(call67); 
        	updateSlides.add(call68); 
        	updateSlides.add(call69); 
        	updateSlides.add(call70); 
        	updateSlides.add(call71); 
        	updateSlides.add(call72); 
        	updateSlides.add(call73); 
        	updateSlides.add(call74); 
        	updateSlides.add(call75); 
        	updateSlides.add(call76); 
        	updateSlides.add(call77); 
        	updateSlides.add(call78); 
        	updateSlides.add(call79); 
        	updateSlides.add(call80); 
        	updateSlides.add(call81); 
        	updateSlides.add(call82); 
        	updateSlides.add(call83); 
        	updateSlides.add(call84); 
        	updateSlides.add(call85); 
        	updateSlides.add(call86);
        	pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides); 
    	} catch(Exception e1) {}

          try {
              try {	 
              	File directoryq = new File(downloadFilepath);
                  File[] filesq = directoryq.listFiles();
                  for (File fq : filesq)
                  {
                      if (fq.getName().endsWith("Citing-Journal-Data-2021.xlsx"))
                      {
                    	  String oldfilename = fq.getName(); 
                      	File myfile = new File(downloadFilepath + oldfilename);
                      	myfile.renameTo(new File(downloadFilepath + "Citing.xlsx"));
                      }
                  }
} catch(Exception e1) {}
              
        String xlsxFilecited = downloadFilepath + "Citing.xlsx";
        Workbook wbd1 = new Workbook();
        wbd1.loadFromFile(xlsxFilecited);
        Worksheet sheetd1 = wbd1.getWorksheets().get(0); 
        for (int i1 = 9; i1 < 30; i1++) {
        	String value = sheetd1.getRange().get("B" + i1).getValue();
              sheetd1.getCellRange("B" + i1).setValue(value);
             if (!value.contentEquals("N/A")) {
//            	 sheetd1.getCellRange("B" + i1).setValue(value);
             }else {
            	 sheetd1.getCellRange("B" + i1).setValue("");
             }
             }
        sheetd1.getRange().get("B9:B50").setNumberFormat("0.000");
        int aa = 13;
        int ab = 9; 
        List<UpdateSlidePojo> updateSlides = new ArrayList<>(); 
      	UpdateSlidePojo call1 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call2 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call3 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call4 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call5 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call6 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call7 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call8 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call9 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call10 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call11 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call12 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call13 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call14 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call15 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call16 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call17 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call18 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call19 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call20 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("B" + ab).getDisplayedText());aa++;ab++;
      	aa = 33;
          ab = 7;
      	UpdateSlidePojo call21 = new UpdateSlidePojo(68,"<tck" + aa + ">","ALL Journals");aa++;ab++;
      	UpdateSlidePojo call22 = new UpdateSlidePojo(68,"<tck" + aa + ">","ALL OTHERS");aa++;ab++;
      	UpdateSlidePojo call23 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call24 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call25 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call26 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call27 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call28 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call29 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call30 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call31 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call32 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call33 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call34 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call35 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call36 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call37 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call38 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call39 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call40 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call41 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call42 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("C" + ab).getDisplayedText());aa++;ab++;
      	aa = 55;
          ab = 7;
      	UpdateSlidePojo call43 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call44 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call45 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call46 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call47 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call48 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call49 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call50 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call51 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call52 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call53 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call54 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call55 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call56 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call57 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call58 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call59 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call60 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call61 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call62 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call63 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call64 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("D" + ab).getDisplayedText());aa++;ab++;
      	aa = 77;
          ab = 7;
      	UpdateSlidePojo call65 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call66 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call67 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call68 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call69 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call70 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call71 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call72 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call73 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call74 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call75 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call76 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call77 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call78 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call79 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call80 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call81 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call82 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call83 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call84 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call85 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
      	UpdateSlidePojo call86 = new UpdateSlidePojo(68,"<tck" + aa + ">",sheetd1.getRange().get("E" + ab).getDisplayedText());aa++;ab++;
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
      	updateSlides.add(call22); 
      	updateSlides.add(call23); 
      	updateSlides.add(call24); 
      	updateSlides.add(call25); 
      	updateSlides.add(call26); 
      	updateSlides.add(call27); 
      	updateSlides.add(call28); 
      	updateSlides.add(call29); 
      	updateSlides.add(call30); 
      	updateSlides.add(call31); 
      	updateSlides.add(call32); 
      	updateSlides.add(call33); 
      	updateSlides.add(call34); 
      	updateSlides.add(call35); 
      	updateSlides.add(call36); 
      	updateSlides.add(call37); 
      	updateSlides.add(call38); 
      	updateSlides.add(call39); 
      	updateSlides.add(call40);     
      	updateSlides.add(call41); 
      	updateSlides.add(call42); 
      	updateSlides.add(call43); 
      	updateSlides.add(call44); 
      	updateSlides.add(call45); 
      	updateSlides.add(call46); 
      	updateSlides.add(call47); 
      	updateSlides.add(call48); 
      	updateSlides.add(call49); 
      	updateSlides.add(call50); 
      	updateSlides.add(call51); 
      	updateSlides.add(call52); 
      	updateSlides.add(call53); 
      	updateSlides.add(call54); 
      	updateSlides.add(call55); 
      	updateSlides.add(call56); 
      	updateSlides.add(call57); 
      	updateSlides.add(call58); 
      	updateSlides.add(call59); 
      	updateSlides.add(call60); 
      	updateSlides.add(call61); 
      	updateSlides.add(call62); 
      	updateSlides.add(call63); 
      	updateSlides.add(call64); 
      	updateSlides.add(call65); 
      	updateSlides.add(call66); 
      	updateSlides.add(call67); 
      	updateSlides.add(call68); 
      	updateSlides.add(call69); 
      	updateSlides.add(call70); 
      	updateSlides.add(call71); 
      	updateSlides.add(call72); 
      	updateSlides.add(call73); 
      	updateSlides.add(call74); 
      	updateSlides.add(call75); 
      	updateSlides.add(call76); 
      	updateSlides.add(call77); 
      	updateSlides.add(call78); 
      	updateSlides.add(call79); 
      	updateSlides.add(call80); 
      	updateSlides.add(call81); 
      	updateSlides.add(call82); 
      	updateSlides.add(call83); 
      	updateSlides.add(call84); 
      	updateSlides.add(call85); 
      	updateSlides.add(call86);
      	pSlide.updatePPTMultipleSlideData(pptid, service, updateSlides);
        } catch(Exception e1) {}}
    	return null;
    }


}