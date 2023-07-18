//Author: Raja
//Date: 28/06/2023

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

public class Order_Excel_xml {
    public static void main(String[] args) {

            // Load Excel file
            InputStream inputStream = Order_Excel_xml.class.getClassLoader().getResourceAsStream("OrderXML.xlsx");
            try {
            //open workbook
                assert inputStream != null;
                Workbook workbook = new XSSFWorkbook(inputStream);

            // Read the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Create XML StringBuilder to store the output
            StringBuilder xmlBuilder = new StringBuilder();
            xmlBuilder.append("<Orders>\n");

            // Iterate through each row
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                // Read the cell values
                String orderNo = row.getCell(0).getStringCellValue();
                double orderTotal = row.getCell(1).getNumericCellValue();
                String orderType = row.getCell(2).getStringCellValue();
                String paymentType = row.getCell(3).getStringCellValue();
                String paymentStatus = row.getCell(4).getStringCellValue();
                String orderStatus = row.getCell(5).getStringCellValue();

                // Create XML element for the current row
                xmlBuilder.append("    <Order>\n");
                xmlBuilder.append("        <OrderNo>").append(orderNo).append("</OrderNo>\n");
                xmlBuilder.append("        <OrderTotal>").append(orderTotal).append("</OrderTotal>\n");
                xmlBuilder.append("        <OrderType>").append(orderType).append("</OrderType>\n");
                xmlBuilder.append("        <PaymentType>").append(paymentType).append("</PaymentType>\n");
                xmlBuilder.append("        <PaymentStatus>").append(paymentStatus).append("</PaymentStatus>\n");
                xmlBuilder.append("        <OrderStatus>").append(orderStatus).append("</OrderStatus>\n");
                xmlBuilder.append("    </Order>\n");
            }

            xmlBuilder.append("</Orders>");

            // Write XML output to a file
            File outputFile = new File("output.xml");
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            outputStream.write(xmlBuilder.toString().getBytes());
            outputStream.close();

            System.out.println("Excel file successfully converted to XML.");
            // Print XML output to console
            System.out.println(xmlBuilder.toString());

            // Close the workbook
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
