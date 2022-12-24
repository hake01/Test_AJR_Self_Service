package com.sn.ajr;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
//import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationRequest;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationResponse;
import com.google.api.services.slides.v1.model.Page;
import com.google.api.services.slides.v1.model.Presentation;
import com.google.api.services.slides.v1.model.ReplaceAllTextRequest;
import com.google.api.services.slides.v1.model.Request;
import com.google.api.services.slides.v1.model.SubstringMatchCriteria;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

public class ProjGSlide {
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    private static final List<String> SCOPES = Collections.singletonList(SlidesScopes.PRESENTATIONS_READONLY);
    private static final String CREDENTIALS_FILE_PATH = "/client_secret.json";

    private Presentation presentation;

    private static Credential authorize(final NetHttpTransport HTTP_TRANSPORT) throws IOException, GeneralSecurityException {
        // Load client secrets.
        InputStream in = ProjGSlide.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		List<String> scopes = Arrays.asList(SheetsScopes.DRIVE);
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
				clientSecrets, scopes)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(
				        System.getProperty("user.home"), ".credentials/2/sheets.googleapis.com-java-quickstart.json")))
				.setAccessType("offline")
				.build();
		
		
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		return credential;
    }
    public Slides getSlideService(String appName) throws IOException , GeneralSecurityException{
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

		Credential credential = authorize(HTTP_TRANSPORT);
		Slides service = new Slides.Builder(HTTP_TRANSPORT, JSON_FACTORY, authorize(HTTP_TRANSPORT))
                .setApplicationName(appName)
                .build();
		return service;
	}
    
    public void updatePresentationData(String presentationId, Slides service,String oldText,String newText) throws IOException {
    	Presentation response = service.presentations().get(presentationId).execute();
        
        List<Page> slides = response.getSlides();
        for (int i = 0; i < slides.size(); ++i) {
        }
        try {
        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
        		.setReplaceAllText(new ReplaceAllTextRequest()
                        .setContainsText(new SubstringMatchCriteria()
                                .setText(oldText)
                                .setMatchCase(true))
                        .setReplaceText(newText)));
        
        BatchUpdatePresentationRequest body =
                new BatchUpdatePresentationRequest().setRequests(requests);
        BatchUpdatePresentationResponse response1 = service.presentations().batchUpdate(presentationId, body).execute();
    } catch(Exception e1) {}
    }
    
    public void updatePPTMultipleSlideData(String presentationId, Slides service,List<UpdateSlidePojo> updateSlides) throws IOException {
//    	System.out.println("Check ");
    	if(Objects.isNull(this.presentation) || (!presentationId.equalsIgnoreCase(this.presentation.getPresentationId()))) {
    		this.presentation = service.presentations().get(presentationId).execute();
    	}
    	try {
            List<Request> requests = new ArrayList<>();
        final List<Page> slides = this.presentation.getSlides();
        
        for(UpdateSlidePojo updateSlide: updateSlides) {
        	int slideNo = updateSlide.getSlideNo() -1;
        	String pageObjectId=slides.get(slideNo).getObjectId();
        	updateSlide.setPageObjectId(pageObjectId);
        	requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest().setPageObjectIds(Arrays.asList(pageObjectId))
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText(updateSlide.getOldText())
                                    .setMatchCase(true))
                            .setReplaceText(updateSlide.getNewText())));
//        	System.out.println("done ");
        }
        BatchUpdatePresentationRequest body =
                new BatchUpdatePresentationRequest().setRequests(requests);
        BatchUpdatePresentationResponse response1 = service.presentations().batchUpdate(presentationId, body).execute();
    } catch(Exception e1) {}
    }

    
    
  }
