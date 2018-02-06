package com.coralnode;

/**
 * Created by Rouzbeh on 31/01/2018.
 */

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;

import  java.io.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFRow;


public class ConvertXMLtoXLS {
    public static void main(String argv[]) {

        try {

            String xlsFileName = "xls/output.xls" ;
            FileOutputStream fileOut = new FileOutputStream(xlsFileName);

            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("FirstSheet");



            File fXmlFile = new File("xml/input.xml");
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(fXmlFile);
            doc.getDocumentElement().normalize();

            System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
            System.out.println("--------");

            NodeList nList = doc.getElementsByTagName("row");
            System.out.println("Number of table rows incl.  header: " + nList.getLength());
            HSSFRow row;

            for (int temp = 0; temp < nList.getLength(); temp++) {

                Node nNode = nList.item(temp);

                if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element eElement = (Element) nNode;
                    System.out.print("row id : ");
                    System.out.println(eElement.getAttribute("id"));
                    NodeList cellNameList = eElement.getElementsByTagName("cell");
                    row = sheet.createRow((short)temp);
                    for (int count = 0; count < cellNameList.getLength(); count++) {
                        Node node1 = cellNameList.item(count);
                        if (node1.getNodeType() == node1.ELEMENT_NODE) {
                            Element cell = (Element) node1;
                            String cellContent = cell.getTextContent();
                            System.out.print(cellContent + " ");
                            row.createCell(count).setCellValue(cellContent);
                        }

                    }
                }
            }
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();

        }
    }
}