import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class OrderLine_Excel_xml {
    public static void main(String[] args) {
        //path to input file
        InputStream inputStream = OrderLine_Excel_xml.class.getResourceAsStream("OrderLineTag.xlsx");
        //open workbook
        try {
            Workbook workbook = new XSSFWorkbook(inputStream);

            //Read the first sheet
            Sheet sheet = workbook.getSheetAt(0);

//            //create document builder factory
//            DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
//            DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();
//
//            //get document
//            Document document = documentBuilder.newDocument();
//
//            //create array list for the excel values
//            ArrayList<ArrayList<String>> list = new ArrayList<>();
//            Element root = document.createElement("Orders");
//            document.appendChild(root);


            //create xml sting builder to store the output
            StringBuilder xmlBuilder = new StringBuilder();
            xmlBuilder.append("<Orders>\n");

            //Iterate through each row
            for (int i = 1;i< sheet.getLastRowNum();i++){
                Row row = sheet.getRow(i);

                //get the cell values
                String OrderNo = row.getCell(0).getStringCellValue();
                int OrderLineNo = (int) row.getCell(1).getNumericCellValue();
                double Qty = row.getCell(2).getNumericCellValue();
                double LineTotal = row.getCell(3).getNumericCellValue();
                String LineStatus = row.getCell(4).getStringCellValue();

                //create XML element for the current row
                xmlBuilder.append("   <Order>\n");
                xmlBuilder.append("        <OrderNo>").append(OrderNo).append("</OrderNo>\n");
                xmlBuilder.append("        <OrderLineNo>").append(OrderLineNo).append("</OrderLineNo>\n");
                xmlBuilder.append("        <Qty>").append(Qty).append("</Qty>\n");
                xmlBuilder.append("        <LineTotal>").append(LineTotal).append("</LineTotal>\n");
                xmlBuilder.append("        <LineStatus>").append(LineStatus).append("</LineStatus>\n");
                xmlBuilder.append("   </Order>\n");
            }
            xmlBuilder.append("</Orders>\n");

            //write XML output to the file
            File outputFile = new File("output.xml");
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            outputStream.write(xmlBuilder.toString().getBytes());
            outputStream.close();

            //print output file to console
            System.out.println("Excel file is converted to XML successfully");
            System.out.println(xmlBuilder.toString());

            //close workbook
            workbook.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
