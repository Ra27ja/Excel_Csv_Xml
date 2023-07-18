//Author: Raja
//Date: 28/06/2023

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.io.FileReader;
import java.io.IOException;

public class Csv_xml {
    public static void main(String[] args) {
        String csvFile = "C:\\Users\\gvajrala\\IdeaProjects\\Excel_Csv_Xml\\src\\main\\resources\\OrderXMLCSV.csv";

        try (FileReader reader = new FileReader(csvFile);
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT)) {

            System.out.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            System.out.println("<Orders>");

            boolean isFirstRecord = true;
            for (CSVRecord record : csvParser) {
                if (isFirstRecord) {
                    isFirstRecord = false;
                    continue; // Skip the header row
                }

                System.out.println("\t<Order>");
                System.out.println("\t\t<OrderNo>" + record.get(0).trim() + "</OrderNo>");
                System.out.println("\t\t<OrderTotal>" + record.get(1).trim() + "</OrderTotal>");
                System.out.println("\t\t<OrderType>" + record.get(2).trim() + "</OrderType>");
                System.out.println("\t\t<PaymentType>" + record.get(3).trim() + "</PaymentType>");
                System.out.println("\t\t<PaymentStatus>" + record.get(4).trim() + "</PaymentStatus>");
                System.out.println("\t\t<OrderStatus>" + record.get(5).trim() + "</OrderStatus>");
                System.out.println("\t</Order>");
            }

            System.out.println("</Orders>");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
