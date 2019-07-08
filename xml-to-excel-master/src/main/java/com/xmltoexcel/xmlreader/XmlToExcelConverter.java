package com.xmltoexcel.xmlreader;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

public class XmlToExcelConverter {
    private static Workbook workbook;
    private static int rowNum;

    private final static int SFT_BASE_FILENAME = 0;
    private final static int SFT_USER_NAME = 1;


    public static void main(String[] args) throws Exception {
        getAndReadXml();
    }

    private static void getAndReadXml() throws Exception {
        System.out.println("getAndReadXml");
        /**Changes*/
        Properties properties = new Properties();
        properties.load(new FileReader("xmltoexcel.properties"));        
        //File xmlFile = new File("C:\\Users\\AG5052068\\Documents\\TP_BatchProv_Tamplate.xml");
        File xmlFile = new File(properties.getProperty("xml-file-location")+properties.getProperty("xml-file-name"));
        initXls();
        Sheet sheet = workbook.getSheetAt(0);
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);
        NodeList nList = doc.getElementsByTagName("InboundFiles");
        for (int i = 0; i < nList.getLength(); i++) {
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                NodeList infil = element.getElementsByTagName("InboundFile");
                for (int j = 0; j < infil.getLength(); j++) {
                    Node inb = infil.item(j);
                    if (inb.getNodeType() == Node.ELEMENT_NODE) {
                        Element sftib = (Element) inb;
                                         
                        String baseFileName = sftib.getElementsByTagName("BaseFileName").item(0).getTextContent();
                        String sftUser = sftib.getElementsByTagName("SFTUserName").item(0).getTextContent();

                        Row row = sheet.createRow(rowNum++);
                        Cell cell = row.createCell(SFT_BASE_FILENAME);
                        cell.setCellValue(baseFileName);

                        cell = row.createCell(SFT_USER_NAME);
                        cell.setCellValue(sftUser);
                    }
                }
            }
        }
        String fileName = new SimpleDateFormat("yyyyMMddHHmm'.xlsx'").format(new Date());
        fileName = properties.getProperty("excel-file-name")+fileName;
        //FileOutputStream fileOut = new FileOutputStream("C:\\Users\\AG5052068\\Documents\\"+fileName);
        FileOutputStream fileOut = new FileOutputStream(properties.getProperty("excel-destination")+fileName);
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

        System.out.println("Excel has been generated..!");
    }


    /**
     * Initializes the POI workbook and writes the header row
     */
    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        //style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(SFT_BASE_FILENAME);
        cell.setCellValue("File Name");
        cell.setCellStyle(style);

        cell = row.createCell(SFT_USER_NAME);
        cell.setCellValue("SFT User Name");
        cell.setCellStyle(style);
    }
}