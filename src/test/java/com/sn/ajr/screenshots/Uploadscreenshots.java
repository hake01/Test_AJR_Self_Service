package com.sn.ajr.screenshots;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.FileContent;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.sn.ajr.ProjGSheet; 
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

import org.testng.annotations.Test;

//import javax.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletRequest;
public class Uploadscreenshots {
	private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
	private static final String TOKENS_DIRECTORY_PATH = "tokens";
	private static final List<String> SCOPES = Collections.singletonList(DriveScopes.DRIVE_FILE);
	private static final String CREDENTIALS_FILE_PATH = "/client_secret.json";
	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
		InputStream in = Uploadscreenshots.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
		if (in == null) {
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline")
				.build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("AJR_Self_Services");
//		Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("bhushan.hake");
//		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		return credential;
	}
	@Test
//	public static String uploadfilebyid(HttpServletRequest req) throws InterruptedException, IOException, GeneralSecurityException{
	public static String uploadfilebyid(HttpServletRequest req) throws InterruptedException, IOException, GeneralSecurityException{ 
		String workingdir = System.getProperty("user.dir");
		String downloadFilepath = workingdir + "\\DownloadingTempfiles\\";
		String server = "test";
		String sheetidq = "1sVBm775JRdpx0hIf8UOmQ4_a8Lxb1r0UXyaewTif8cc";//live
		if(server.contentEquals("test")) {
		sheetidq = "1qWzW6n27NTUb-tsXdKMDuWkO0zl43YmbV1GTB-JbGyY"; //test
		}
		ProjGSheet sheetq = new ProjGSheet();
		Sheets sheetsService1 = sheetq.getSheetsService("Google Sheets Example");   
	  	String sheetname = "";
	  	try {
	  	List<List<Object>> values = sheetq.getCellData(sheetsService1, "queuecheck!D2:N2",sheetidq); 
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
		try {
		String filename0 = "Cover.png"; 
		String filename1 = "MT_1.png";
		String filename2 = "MT_2.png";
		String filename3 = "SJR.png";
		String filename4 = "5YearIF.png";
		String filename5 = "impact-factor.png";
		String filename6 = "newplot.png";
		String fileformat = "image/jpeg";		
		String fileid0 = uploadfile(downloadFilepath, filename0, fileformat);
		String fileid1 = uploadfile(downloadFilepath, filename1, fileformat);
		String fileid2 = uploadfile(downloadFilepath, filename2, fileformat);
		String fileid3 = uploadfile(downloadFilepath, filename3, fileformat);
		String fileid4 = uploadfile(downloadFilepath, filename4, fileformat);
		String fileid5 = uploadfile(downloadFilepath, filename5, fileformat);
		String fileid6 = uploadfile(downloadFilepath, filename6, fileformat);
		sheet.updateCellvalue(sheetsService, sheetname + "!L4", "AuploadedfilesA: |" + fileid0 + "|" + fileid1 + "|" + fileid2 + "|" + fileid3 + "|" + fileid4 + "|" + fileid5 + "|" + fileid6,sheetid);
		Thread.sleep(1000);
} catch(Exception e1) {}} 
}
		return sheetname;
	}
	
	public static String uploadfile(String downloadFilepath, String filename,String fileformat) throws IOException, GeneralSecurityException {
		String fileid = "";
		System.out.println("uploadfile Check" + downloadFilepath);
		try {
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			Drive service = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT)).build();
			File fileMetadata = new File();
			fileMetadata.setName(filename);
			System.out.println("uploadfile Check" + filename);
			java.io.File filePath = new java.io.File(downloadFilepath + filename);
			FileContent mediaContent = new FileContent(fileformat, filePath);
			File file = service.files().create(fileMetadata, mediaContent)
					.setFields("id")
					.execute();        
			fileid = filename + "^" + file.getId();
		} catch(Exception e1) {}
		return fileid;
	}
	 

}
