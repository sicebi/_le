package com.company;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

public class Letter {

    public void replaceInfo(String letterURL){
        //, Map<String,String> details, List<Map<String, String>> transactions, String base){
        String CompletedFileURL = "C:\\Users\\r0000382\\Desktop\\Completed latter\\CompletedLetter.docx";


        InputStream fis = null;
        try {
            fis = new FileInputStream(letterURL);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XWPFDocument document = null;
        try {
            document = new XWPFDocument(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
//                getting information from the table
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            for (XWPFTableRow row : table.getRows()) {

                for (XWPFTableCell cell : row.getTableCells()) {
//                    System.out.println(">>>>>>>" + cell.getText());
                    for (XWPFParagraph tableP : cell.getParagraphs())
                        for (XWPFRun run : tableP.getRuns()) {
                            String text = run.getText(0);
//                            System.out.println(">>>>>>>" + cell.getText());
//                            System.out.println(">>>>>>>" + text);

                            if (text.contains("DateRep")) {
                                text = text.replace("DateRep", "2020-03-06");
                            }
                            if (text.contains("ReferenceRep")) {
                                text = text.replace("ReferenceRep", "123456789");
                            }
                            if (text.contains("TotalfraudRep")) {
                                text = text.replace("TotalfraudRep", "3000000");
//                                run.setText(text,0);
                            }
                            if (text.contains("TotalrefundRep")) {
                                text = text.replace("TotalrefundRep", "90000");
//                                run.setText(text,0);
                            }
                            if (text.contains("TotallossRep")) {
                                text = text.replace("TotallossRep", "29999888.00");
//                                run.setText(text,0);
                            }

                            run.setText(text, 0);
                            System.out.println("cell text------ ++++++ " + text);

                        }
                }

                System.out.println(" ");
            }}


//        getting the paragraphs only
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            String text = paragraphs.get(i).getText();
            List<XWPFRun> runs = paragraphs.get(i).getRuns();
            System.out.println("Paragraph " + i + ": " + text);

            for (int j = 0; j < runs.size(); j++ ){
                String text2 = runs.get(j).getText(0);

                if (text2 != null){
                    text2 = text2.replace("NameRep", "Advocate");
                }
                if (text2 != null){
                    text2 = text2.replace("InitialsRep", "SA");
                }
                if (text2 != null){
                    text2 = text2.replace("SurnameRep", "Ntini");
                }
                if (text2 != null){
                    text2 = text2.replace("AddressLineRep", "10111 street");
                }
                if (text2 != null){
                    text2 = text2.replace("CityRep", "Johannesburg");
                }

                if (text2 != null){
                    text2 = text2.replace("TitleRep", "MR");
                }
                if (text2 != null){
                    text2 = text2.replace("CodeRep", "2000");
                }
                if (text2 != null){
                    text2 = text2.replace("Sequence of events", ".............Paragraph................ ");
                }
                if (text2 != null) {
                    text2 = text2.replace("SignatureRep", "SA Ntini");
                }
                runs.get(j).setText(text2,0);
                System.out.println("    Run " + j + ": " + text2);
            }
        }

        OutputStream out = null;
        try {
            out = new FileOutputStream(CompletedFileURL);
            document.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args){
        Letter test = new Letter();
        test.replaceInfo("C:\\Users\\r0000382\\Desktop\\CRC Letter New.docx");
    }

}
