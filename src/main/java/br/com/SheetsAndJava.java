/*
 * Google Sheets and Java Integration
 *
 * This program demonstrates how to integrate Google Sheets with Java using the Google Sheets API.
 * It allows updating a Google Sheets spreadsheet with calculated values based on certain conditions.
 *
 * Prerequisites:
 * - Google Sheets API enabled for your Google Cloud project
 * - Google Cloud project credentials set up and saved in a 'credentials.json' file
 * - Google Sheets spreadsheet ID obtained from the URL of the spreadsheet
 *
 * Steps:
 * 1. Authorization: The program authorizes itself using OAuth 2.0 with the credentials stored in 'credentials.json'.
 * 2. Accessing Sheets: After authorization, it accesses the Google Sheets API.
 * 3. Updating Sheets: It reads data from a specified range in the spreadsheet, performs calculations,
 *    and updates the spreadsheet with the calculated values.
 *
 * Author: [Magno Luis]
 *
 */

package br.com;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetsAndJava {

    // Google Sheets service instance
    private static com.google.api.services.sheets.v4.Sheets sheetsService;
    // Application name
    private static final String APPLICATION_NAME = "Integrating Google Spreadsheet with Java";
    // Spreadsheet ID (replace with your own)
    private static final String SPREADSHEET_ID = "14umGRUJ3cWCMU0wnz4xm6K6ZGEYuTkNnhTBXFJu3xfU";
    // JSON factory
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    // Maximum allowed number of fouls
    private static final int MAXIMUM_FOULS = 15;

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        sheetsService = getSheetsService();
        // Specify the range of the spreadsheet to update
        String range = "engenharia_de_software!A4:F27";
        updateSheet(range);
    }

    // Authorize the application using OAuth 2.0
    private static Credential authorize() throws IOException, GeneralSecurityException {
        InputStream in = SheetsAndJava.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
                .setAccessType("offline")
                .build();

        return new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
    }

    // Get the Google Sheets service instance
    private static com.google.api.services.sheets.v4.Sheets getSheetsService() throws IOException, GeneralSecurityException {
        Credential credential = authorize();
        return new com.google.api.services.sheets.v4.Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    // Format a double value to one decimal place
    private static double format(double n) {
         int intPart = (int) n;
         double decimalPart = n - intPart;

         double roundDecimal = Math.round(decimalPart *  10) / 10.0;
         return (roundDecimal >= 0.5) ? Math.ceil(n * 10) / 10.0 : Math.floor(n * 10) / 10.0;
    }

    // Calculate the average of three integers
    private static double calculateAverage(int note1, int note2, int note3) {
        return (note1 + note2 + note3) / 30.0;
    }

    // Determine the situation based on fouls and average
    private static String determineSituation(int fouls, double average) {
        return (fouls > MAXIMUM_FOULS) ? "Reprovado por Falta" :
                (average < 5.0) ? "Reprovado por Nota" :
                (average >= 5.0 && average < 7.0) ? "Exame Final" : "Aprovado";
    }

    // Calculate the NAF (Nota para Aprovação Final)
    private static double calculateNAF(double average) {
        return 10 - average;
    }

    // Update the specified range in the spreadsheet
    private static void updateSheet(String range) throws IOException {
        ValueRange response = sheetsService.spreadsheets().values().get(SPREADSHEET_ID, range).execute();
        List<List<Object>> values = response.getValues();

        for (List<Object> row : values) {
            int fouls = Integer.parseInt((String) row.get(2));
            int note1 = Integer.parseInt((String) row.get(3));
            int note2 = Integer.parseInt((String) row.get(4));
            int note3 = Integer.parseInt((String) row.get(5));
            double average = calculateAverage(note1, note2, note3);
            double formattedAverage = format(average);
            System.out.println(formattedAverage);
            String situation = determineSituation(fouls, average);
            double naf = 0.0;

            naf = situation.equals("Exame Final") ? format(calculateNAF(formattedAverage)) : naf;

            updateCellValue(values.indexOf(row) + 4, "G", situation);
            updateCellValue(values.indexOf(row) + 4, "H", situation.equals("Exame Final") ? naf : 0);
        }
    }

    // Update cell value
    private static void updateCellValue(int rowIndex, String column, Object value) throws IOException {
        String cellRange = column + rowIndex;
        List<Object> data = new ArrayList<>();
        data.add(value);
        ValueRange body = new ValueRange().setValues(Arrays.asList(data));
        sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, cellRange, body)
                .setValueInputOption("RAW")
                .execute();
    }
}

