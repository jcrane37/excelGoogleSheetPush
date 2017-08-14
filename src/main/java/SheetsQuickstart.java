package main.java;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class SheetsQuickstart {
    /** Application name. */
    private static final String APPLICATION_NAME =
            "Google Sheets API Java Quickstart";

    /** Directory to store user credentials for this application. */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(
            System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart.json");

    /** Global instance of the {@link FileDataStoreFactory}. */
    private static FileDataStoreFactory DATA_STORE_FACTORY;

    /** Global instance of the JSON factory. */
    private static final JsonFactory JSON_FACTORY =
            JacksonFactory.getDefaultInstance();

    /** Global instance of the HTTP transport. */
    private static HttpTransport HTTP_TRANSPORT;

    /** Global instance of the scopes required by this quickstart.
     *
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/sheets.googleapis.com-java-quickstart.json
     */
    private static final List<String> SCOPES =
            Arrays.asList(SheetsScopes.SPREADSHEETS);
    //path to excel document
    private static final String FILE_NAME = "/Users/josephcrane/Downloads/GoogleApiTest3/GoogleApiTest2/src/main/Electricity_Sample_Big_EDIC_File.xlsx";
    //ID for google sheets
    public static final String spreadsheetId = "1Kwmu0_DZh7KAHDY-iaV99b0_UV6P3YZpGw7r7j9-GSU";


    //convert column number to excel row number
    public static String convertToNumber(int n) {
        if(n <= 0){
            throw new IllegalArgumentException("Input is not valid!");
        }

        StringBuilder sb = new StringBuilder();

        while(n > 0){
            n--;
            char ch = (char) (n % 26 + 'A');
            n /= 26;
            sb.append(ch);
        }

        sb.reverse();
        return sb.toString();
    }

    public static ValueRange response;
    public static UpdateValuesResponse request;

    public static void main (String[] args) throws Exception {
        List<List<Object>> arrData = getData();
        //List<List<Object>> values = SheetsQuickstart.getResponse("BrowserSheet","A1","A").getValues ();
        SheetsQuickstart.setValue("Sheet1","A1",convertToNumber(columns), arrData);
    }


    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }


    /**
     * Creates an authorized Credential object.
     * @return an authorized Credential object.
     * @throws IOException
     */
    public static Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in =
                SheetsQuickstart.class.getResourceAsStream("../resources/client_secret.json");
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Build and return an authorized Sheets API client service.
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    public static Sheets getSheetsService() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static ValueRange getResponse(String SheetName,String RowStart, String RowEnd) throws IOException{
        // Build a new authorized API client service.
        Sheets service = getSheetsService();


        // Prints the names and majors of students in a sample spreadsheet:
        //String spreadsheetId = "1djmJ_n6T4vCE3rviomC1kKVKpx0eCCfzktnNk4rGQ4c";
        String range = SheetName+"!"+RowStart+":"+RowEnd;
        response = service.spreadsheets().values()
                .get(spreadsheetId, range).execute ();

        return response;

    }


    public static void setValue(String SheetName,String RowStart, String RowEnd, List<List<Object>> arrData) throws IOException{
        // Build a new authorized API client service.
        Sheets service = getSheetsService();
        // Prints the names and majors of students in a sample spreadsheet:
        //String spreadsheetId = "1djmJ_n6T4vCE3rviomC1kKVKpx0eCCfzktnNk4rGQ4c";
        String range = RowStart+":"+RowEnd;



        ValueRange oRange = new ValueRange();
        oRange.setRange(range); // I NEED THE NUMBER OF THE LAST ROW
        oRange.setValues(arrData);

        List<ValueRange> oList = new ArrayList<>();
        oList.add(oRange);

        BatchUpdateValuesRequest oRequest = new BatchUpdateValuesRequest();
        oRequest.setValueInputOption("RAW");
        oRequest.setData(oList);

        BatchUpdateValuesResponse oResp1 = service.spreadsheets().values().batchUpdate(spreadsheetId, oRequest).execute();

        // service.spreadsheets().values().update (spreadsheetId, range,) ;
        //return request;

    }

    public static int columns = 0;

    public static List<List<Object>> data = new ArrayList<List<Object>>();

    public static List<List<Object>> getData ()  {

        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);


            int rows = 0;
            ArrayList<Integer> numbers = new ArrayList<Integer>();
            BufferedReader console = new BufferedReader(new InputStreamReader(System.in));



            Iterator<Row> iterator = datatypeSheet.iterator();
            int count = 0;

            while (iterator.hasNext()) {
                List<Object> data1 = new ArrayList<>();

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();


                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0

                    if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                        data1.add(currentCell.getStringCellValue());
                    } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        data1.add(currentCell.getNumericCellValue());
                    } else if (currentCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        data1.add(currentCell.getStringCellValue());
                    }

                    if(count == 0) {

                        columns += 1;
                    }

                }
                count += 1;

                rows += 1;

                data.add(data1);

            }
        } catch (IOException e) {
        e.printStackTrace();
        }

        return data;

    }

//for personal use
/*
    public static List<List<Object>> getData2 ()  {

        List<Object> data1 = new ArrayList<Object>();
        data1.add ("Ashwin");
        data1.add ("bacon");
        data1.add ("meatball");

        //System.out.println(data);

        List<List<Object>> data = new ArrayList<List<Object>>();
        data.add (data1);

        System.out.println(data);

        return data;
    } */


}