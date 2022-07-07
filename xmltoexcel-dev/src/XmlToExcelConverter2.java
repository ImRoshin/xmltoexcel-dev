//import org.apache.commons.io.FileUtils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
//import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Scanner;
//import java.net.URL;

public class XmlToExcelConverter2 {


    public static void main(String[] args) throws Exception {

        Scanner obj = new Scanner(System.in);
        System.out.println("Enter number of new lines : ");

        int n = obj.nextInt();
        obj.close();
        getAndReadXml(n);
    }
    private static Workbook workbook;
    private static int rowNum;
    private final static int CardNumberRange = 0;
    private final static int CardStatus= 1;
    private final static int Fields = 2;
    private final static int FileType = 3;
    private final static int TransactionCode = 4;
    private final static int AccountNumber = 5;
    private final static int AcquirerReferenceNumber = 6;
    private final static int AcquirersBusinessID= 7;
    private final static int PurchaseDate = 8;
    private final static int SourceAmount = 9;
    private final static int SourceCurrencyCode = 10;
    private final static int DestinationAmount = 11;
    private final static int DestinationCurrencyCode = 12;
    private final static int MerchantCountryCode = 13;
    private final static int MerchantName = 14;
    private final static int MerchantCity = 15;
    private final static int MerchantZIP = 16;
    private final static int MerchantCategoryCode = 17;
    private final static int RequestedPaymentService = 18;
    private final static int UsageCode = 19;
    private final static int ReasonCode = 20;
    private final static int SettlementFlag = 21;
    private final static int AuthorizationCharacteristicsIndicator = 22;
    private final static int AuthorizationCode = 23;
    private final static int POSTerminalCapability = 24;
    private final static int CardholderIDMethod = 25;
    private final static int POSEntryMode = 26;
    private final static int CentralProcessingDate = 27;
    private final static int CardAcceptorID = 28;
    private final static int TerminalID = 29;
    private final static int AuthorizedAmount = 30;
    private final static int AuthorizationCurrencyCode = 31;
    private final static int AuthorizationResponseCode = 32;

    private static void getAndReadXml(int n) throws Exception {
        System.out.println("getAndReadXml");

        File xmlFile = new File("C:\\Users\\jkoilpillai\\intellij\\config\\config3.xml");

        initXls();

        Sheet sheet = workbook.getSheetAt(0);

        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);

        while(n>0)
        {
            NodeList nList = doc.getElementsByTagName("row");
            for (int i = 0; i < nList.getLength(); i++) {
                Node node = nList.item(i);
                if (node.getNodeType() == Node.ELEMENT_NODE) {
                    Element element = (Element) node;
                    String Transaction_Code = element.getElementsByTagName("Transaction_Code").item(0).getTextContent();
                    String Account_Number = element.getElementsByTagName("Account_Number").item(0).getTextContent();
                    String Acquirer_Reference_Number = element.getElementsByTagName("Acquirer_Reference_Number").item(0).getTextContent();
                    String Acquirers_Business_ID = element.getElementsByTagName("Acquirers_Business_ID").item(0).getTextContent();
                    String Authorization_Characteristics_Indicator = element.getElementsByTagName("Authorization_Characteristics_Indicator").item(0).getTextContent();
                    String Authorization_Code = element.getElementsByTagName("Authorization_Code").item(0).getTextContent();
                    String Authorization_Currency_Code = element.getElementsByTagName("Authorization_Currency_Code").item(0).getTextContent();
                    String Authorization_Response_Code = element.getElementsByTagName("Authorization_Response_Code").item(0).getTextContent();
                    String Authorized_Amount = element.getElementsByTagName("Authorized_Amount").item(0).getTextContent();
                    String CardNumber_Range = element.getElementsByTagName("CardNumber_Range").item(0).getTextContent();
                    String Card_Status = element.getElementsByTagName("Card_Status").item(0).getTextContent();
                    String Card_Acceptor_ID = element.getElementsByTagName("Card_Acceptor_ID").item(0).getTextContent();
                    String Cardholder_ID_Method = element.getElementsByTagName("Cardholder_ID_Method").item(0).getTextContent();
                    String Central_Processing_Date_YDDD = element.getElementsByTagName("Central_Processing_Date_YDDD").item(0).getTextContent();
                    //String Destination_Amount = element.getElementsByTagName("Destination_Amount").item(0).getTextContent();
                    String Fields1 = element.getElementsByTagName("Fields").item(0).getTextContent();
                    String FileType1 = element.getElementsByTagName("FileType").item(0).getTextContent();
                    String Merchant_Category_Code = element.getElementsByTagName("Merchant_Category_Code").item(0).getTextContent();
                    String Merchant_Country_Code = "US", Source_Currency_Code = "840", Destination_Currency_Code = "840";
                    String Merchant_City = element.getElementsByTagName("Merchant_City").item(0).getTextContent();
                    String Merchant_Name = element.getElementsByTagName("Merchant_Name").item(0).getTextContent();
                    String Merchant_ZIP_Code = element.getElementsByTagName("Merchant_ZIP_Code").item(0).getTextContent();
                    String POS_Entry_Mode = element.getElementsByTagName("POS_Entry_Mode").item(0).getTextContent();
                    String POS_Terminal_Capability = element.getElementsByTagName("POS_Terminal_Capability").item(0).getTextContent();
                    String Purchase_Date_MMDD = element.getElementsByTagName("Purchase_Date_MMDD").item(0).getTextContent();
                    String Reason_Code = element.getElementsByTagName("Reason_Code").item(0).getTextContent();
                    String Requested_Payment_Service = element.getElementsByTagName("Requested_Payment_Service").item(0).getTextContent();
                    String Settlement_Flag = element.getElementsByTagName("Settlement_Flag").item(0).getTextContent();
                    String Source_Amount = element.getElementsByTagName("Source_Amount").item(0).getTextContent();
                    String Terminal_ID = element.getElementsByTagName("Terminal_ID").item(0).getTextContent();
                    String Usage_Code = element.getElementsByTagName("Usage_Code").item(0).getTextContent();

                    NodeList Merchant = element.getElementsByTagName("Merchant");
                    for (int j = 0; j < Merchant.getLength(); j++) {
                        Node merch = Merchant.item(j);
                        if (merch.getNodeType() == Node.ELEMENT_NODE) {
                            Element product = (Element) merch;
                            Merchant_Country_Code = product.getElementsByTagName("Merchant_Country_Code").item(0).getTextContent();
                            Source_Currency_Code = product.getElementsByTagName("Source_Currency_Code").item(0).getTextContent();
                            Destination_Currency_Code = product.getElementsByTagName("Destination_Currency_Code").item(0).getTextContent();
                        }

                        Row row = sheet.createRow(rowNum++);
                        Cell cell;
                        cell = row.createCell(CardNumberRange);
                        cell.setCellValue(CardNumber_Range);

                        cell = row.createCell(CardStatus);
                        cell.setCellValue(Card_Status);

                        cell = row.createCell(Fields);
                        cell.setCellValue(Fields1);

                        cell = row.createCell(FileType);
                        cell.setCellValue(FileType1);

                        cell = row.createCell(TransactionCode);
                        cell.setCellValue(Transaction_Code);

                        cell = row.createCell(AccountNumber);
                        cell.setCellValue(Account_Number);

                        cell = row.createCell(AcquirerReferenceNumber);
                        cell.setCellValue(Acquirer_Reference_Number);

                        cell = row.createCell(AcquirersBusinessID);
                        cell.setCellValue(Acquirers_Business_ID);

                        cell = row.createCell(PurchaseDate);
                        cell.setCellValue(Purchase_Date_MMDD);

                        cell = row.createCell(SourceAmount);
                        cell.setCellValue(Source_Amount);

                        cell = row.createCell(SourceCurrencyCode);
                        cell.setCellValue(Source_Currency_Code);

                        cell = row.createCell(DestinationAmount);
                        cell.setCellValue(Source_Amount);

                        cell = row.createCell(DestinationCurrencyCode);
                        cell.setCellValue(Destination_Currency_Code);

                        cell = row.createCell(MerchantCountryCode);
                        cell.setCellValue(Merchant_Country_Code);

                        cell = row.createCell(MerchantName);
                        cell.setCellValue(Merchant_Name);

                        cell = row.createCell(MerchantCity);
                        cell.setCellValue(Merchant_City);

                        cell = row.createCell(MerchantZIP);
                        cell.setCellValue(Merchant_ZIP_Code);

                        cell = row.createCell(MerchantCategoryCode);
                        cell.setCellValue(Merchant_Category_Code);

                        cell = row.createCell(RequestedPaymentService);
                        cell.setCellValue(Requested_Payment_Service);

                        cell = row.createCell(UsageCode);
                        cell.setCellValue(Usage_Code);

                        cell = row.createCell(ReasonCode);
                        cell.setCellValue(Reason_Code);

                        cell = row.createCell(SettlementFlag);
                        cell.setCellValue(Settlement_Flag);

                        cell = row.createCell(AuthorizationCharacteristicsIndicator);
                        cell.setCellValue(Authorization_Characteristics_Indicator);

                        cell = row.createCell(AuthorizationCode);
                        cell.setCellValue(Authorization_Code);

                        cell = row.createCell(POSTerminalCapability);
                        cell.setCellValue(POS_Terminal_Capability);

                        cell = row.createCell(CardholderIDMethod);
                        cell.setCellValue(Cardholder_ID_Method);

                        cell = row.createCell(POSEntryMode);
                        cell.setCellValue(POS_Entry_Mode);

                        cell = row.createCell(CentralProcessingDate);
                        cell.setCellValue(Central_Processing_Date_YDDD);

                        cell = row.createCell(CardAcceptorID);
                        cell.setCellValue(Card_Acceptor_ID);

                        cell = row.createCell(TerminalID);
                        cell.setCellValue(Terminal_ID);

                        cell = row.createCell(AuthorizedAmount);
                        cell.setCellValue(Authorized_Amount);

                        cell = row.createCell(AuthorizationCurrencyCode);
                        cell.setCellValue(Authorization_Currency_Code);

                        cell = row.createCell(AuthorizationResponseCode);
                        cell.setCellValue(Authorization_Response_Code);

                    }
                }
            }
            n--;
        }

        sheet.autoSizeColumn(35);
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\jkoilpillai\\intellij\\config\\Clearing_Transactions_VISA_BASE2.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

        System.out.println("getAndReadXml finished, processed " );
    }

    private static void get_list_set_random(String var){
       
        NodeList nList = doc.getElementsByTagName("row");
        for (int i = 0; i < nList.getLength(); i++) {
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                NodeList list = element.getElementsByTagName("Merchant");
            for (int j = 0; j < Merchant.getLength(); j++) {
                Node merch = Merchant.item(j);
                if (merch.getNodeType() == Node.ELEMENT_NODE) {
                    Element product = (Element) merch;
                    Merchant_Country_Code = product.getElementsByTagName("Merchant_Country_Code").item(0).getTextContent();
                    Source_Currency_Code = product.getElementsByTagName("Source_Currency_Code").item(0).getTextContent();
                    Destination_Currency_Code = product.getElementsByTagName("Destination_Currency_Code").item(0).getTextContent();
                }

    }
    private static void initXls() {
        workbook = new XSSFWorkbook();

        System.out.println("Creating excel " );

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        CellStyle stringStyle = workbook.createCellStyle();
        stringStyle.setAlignment(HorizontalAlignment.CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(CardNumberRange);
        cell.setCellValue("Card Number Range");
        cell.setCellStyle(style);

        cell = row.createCell(CardStatus);
        cell.setCellValue("Card Status");
        cell.setCellStyle(style);

        cell = row.createCell(Fields);
        cell.setCellValue("Fields");
        cell.setCellStyle(style);

        cell = row.createCell(FileType);
        cell.setCellValue("File Type");
        cell.setCellStyle(style);

        cell = row.createCell(TransactionCode);
        cell.setCellValue("Transaction Code");
        cell.setCellStyle(style);

        cell = row.createCell(AccountNumber);
        cell.setCellValue("Account Number");
        cell.setCellStyle(style);

        cell = row.createCell(AcquirerReferenceNumber);
        cell.setCellValue("Acquirer Reference Number");
        cell.setCellStyle(style);


        cell = row.createCell(AcquirersBusinessID);
        cell.setCellValue("Acquirers Business ID");
        cell.setCellStyle(style);

        cell = row.createCell(PurchaseDate);
        cell.setCellValue("Purchase Date");
        cell.setCellStyle(style);

        cell = row.createCell(SourceAmount);
        cell.setCellValue("Source Amount");
        cell.setCellStyle(style);

        cell = row.createCell(SourceCurrencyCode);
        cell.setCellValue("Source Currency Code");
        cell.setCellStyle(style);

        cell = row.createCell(DestinationAmount);
        cell.setCellValue("Destination Amount");
        cell.setCellStyle(style);

        cell = row.createCell(DestinationCurrencyCode);
        cell.setCellValue("Destination Currency Code");
        cell.setCellStyle(style);

        cell = row.createCell(MerchantCountryCode);
        cell.setCellValue("Merchant Country Code");
        cell.setCellStyle(style);

        cell= row.createCell(MerchantName);
        cell.setCellValue("Merchant Name");
        cell.setCellStyle(style);

        cell = row.createCell(MerchantCity);
        cell.setCellValue("Merchant City");
        cell.setCellStyle(style);

        cell = row.createCell(MerchantZIP);
        cell.setCellValue("Merchant ZIP code");
        cell.setCellStyle(style);

        cell = row.createCell(MerchantCategoryCode);
        cell.setCellValue("Merchant Category Code");
        cell.setCellStyle(style);

        cell = row.createCell(RequestedPaymentService);
        cell.setCellValue("Requested Payment Service");
        cell.setCellStyle(style);

        cell = row.createCell(UsageCode);
        cell.setCellValue("Usage Code");
        cell.setCellStyle(style);

        cell = row.createCell(ReasonCode);
        cell.setCellValue("Reason Code");
        cell.setCellStyle(style);

        cell = row.createCell(SettlementFlag);
        cell.setCellValue("Settlement Flag");
        cell.setCellStyle(style);

        cell = row.createCell(AuthorizationCharacteristicsIndicator);
        cell.setCellValue("Authorization Characteristics Indicator");
        cell.setCellStyle(style);

        cell = row.createCell(AuthorizationCode);
        cell.setCellValue("Authorization Code");
        cell.setCellStyle(style);

        cell = row.createCell(POSTerminalCapability);
        cell.setCellValue("POS Terminal Capability");
        cell.setCellStyle(style);

        cell = row.createCell(CardholderIDMethod);
        cell.setCellValue("Cardholder ID Method");
        cell.setCellStyle(style);

        cell = row.createCell(POSEntryMode);
        cell.setCellValue("POS EntryMode");
        cell.setCellStyle(style);

        cell = row.createCell(CentralProcessingDate);
        cell.setCellValue("Central Processing Date (YDDD)");
        cell.setCellStyle(style);

        cell = row.createCell(CardAcceptorID);
        cell.setCellValue("Card Acceptor ID");
        cell.setCellStyle(style);

        cell = row.createCell(TerminalID);
        cell.setCellValue("Terminal ID");
        cell.setCellStyle(style);

        cell = row.createCell(AuthorizedAmount);
        cell.setCellValue("Authorized Amount");
        cell.setCellStyle(style);

        cell = row.createCell(AuthorizationCurrencyCode);
        cell.setCellValue("Authorization Currency Code");
        cell.setCellStyle(style);

        cell = row.createCell(AuthorizationResponseCode);
        cell.setCellValue("Authorization Response Code");
        cell.setCellStyle(style);
    }
}
