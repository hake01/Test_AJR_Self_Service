package com.sn.ajr; 
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java .util.List;

import java.util.Arrays;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.sn.ajr.ProjGSheet;
//import jakarta.servlet.*;
//import jakarta.servlet.http.*;
import jakarta.servlet.http.HttpServletRequest;
public class ProjGSheet {
	private Credential authorize() throws IOException, GeneralSecurityException{
		InputStream in = ProjGSheet.class.getResourceAsStream("/client_secret.json");
		System.out.println(JacksonFactory.getDefaultInstance());
		System.out.println("Cheking in" + in);
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(), new InputStreamReader(in));
		List<String> scopes = Arrays.asList(SheetsScopes.DRIVE);
		
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
				clientSecrets, scopes)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(
				        System.getProperty("user.home"), ".credentials/2/sheets.googleapis.com-java-quickstart.json")))
				.setAccessType("offline")
				.build();
		
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("AJR_Self_Services");
//		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		return credential;
	}
	
	public Sheets getSheetsService(String appName) throws IOException , GeneralSecurityException{
		Credential credential = authorize();
		return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(), credential)
				.setApplicationName(appName).build();
	}
	
	public void updateCellvalue(Sheets sheetsService,String cellAddr, String value,String spreadSheetId) throws IOException, GeneralSecurityException {
		System.out.println("getmethod");	
		ValueRange body = new ValueRange().setValues(Arrays.asList(Arrays.asList(value)));
		try {
			System.out.println(spreadSheetId);	
			System.out.println(cellAddr);	
			System.out.println(body);	
		UpdateValuesResponse result = sheetsService.spreadsheets().values()
				.update(spreadSheetId, cellAddr, body)
				.setValueInputOption("Raw")
				.execute();
		System.out.println(result);	
	} catch(Exception e1) {
		System.out.println(e1);	
		System.out.println(e1.getMessage());	
	}
	}
	public List<List<Object>> getCellData(Sheets sheetsService,final String range, String spreadSheetId) throws IOException, GeneralSecurityException {
		System.out.println(spreadSheetId);	
		System.out.println(range);
		com.google.api.services.sheets.v4.model.ValueRange response = sheetsService.spreadsheets().values()
                .get(spreadSheetId, range)
                .execute();
        return response.getValues();
	}
	
}