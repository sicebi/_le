package com.company;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

public class letter {

    List<String> replacementFields;
    Map<String, String> caseLevel;

    public letter() {

        caseLevel = new HashMap<String, String>();
        caseLevel.put("Card Account Takeover", "Identity Takeover");
        caseLevel.put("Card Lost", "Lost");
        caseLevel.put("Card Not Present", "Fraudulent use of card number");
        caseLevel.put("Card Not Received/Intercepted", "Identity Takeover");
        caseLevel.put("Card Stolen", "Stolen");
        caseLevel.put("CardLost", "Lost");
        caseLevel.put("CardStolen", "Stolen");
        caseLevel.put("Counterfeit (POS Skimming)", "Counterfeit");
        caseLevel.put("Counterfeit Card", "Counterfeit");
        caseLevel.put("False Fraudulent Card Application", "Identity Takeover");
        caseLevel.put("Fraudulent Instruction", "Identity Takeover");
        caseLevel.put("Stolen", "Stolen");

        replacementFields = new ArrayList<String>();
        replacementFields.add("FullDate");
        replacementFields.add("TitleRep");
        replacementFields.add("InitialsRep");
        replacementFields.add("SurnameRep");
        replacementFields.add("AddressLineRep");
        replacementFields.add("DistrictRep");
        replacementFields.add("CityRep");
        replacementFields.add("CodeRep");
//        tested here ------------------------------------
        replacementFields.add("DateRep");
        replacementFields.add("ReferenceRep");
        replacementFields.add("SalutationRep");
        replacementFields.add("TotalfraudRep");
        replacementFields.add("TotalrefundRep");
        replacementFields.add("TotallossRep");
//        replacementFields.add("OurRefTitleRep");
//        replacementFields.add("OurRefInitialsRep");
        replacementFields.add("Summary of eventsRep");
   //     replacementFields.add("LapseDateRep");

    }

    public void replaceInfo(String letterURL){
        //, Map<String,String> details, List<Map<String, String>> transactions, String base){
        String CompletedFileURL = "C:\\Users\\A239590\\Desktop\\Completed.docx";


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
                    for (XWPFParagraph tableP : cell.getParagraphs()){
                        for (XWPFRun run : tableP.getRuns()){
                            String text = run.getText(0);
//                            System.out.println(">>>>>>>" + cell.getText());
//                            System.out.println(">>>>>>>" + text);

                            if (text.contains("DateRep")){
                                text = text.replace("DateRep", "2020-03-06");
                            }
                            if (text.contains("ReferenceRep")){
                                text = text.replace("ReferenceRep", "123456789");
                            }
                            if (text.contains("TotalfraudRep")){
                                text = text.replace("TotalfraudRep", "3000000");
                                run.setText(text,0);
                            }
                            if (text.contains("TotalrefundRep")){
                                text = text.replace("TotalrefundRep", "90000");
                                run.setText(text,0);
                            }
                            if (text.contains("TotallossRep")){
                                text = text.replace("TotallossRep", "29999888.00");
                                run.setText(text,0);
                            }

                            run.setText(text,0);
                            System.out.println("cell text------ ++++++ " + text);

                        }
                    }

//                    System.out.println(cell.getText());
//                    String sFieldValue = cell.getText();
//                    if (sFieldValue.matches("Date") || sFieldValue.matches("Approved")) {
//                        System.out.println("The match as per the Document is True");
//                    }
//					System.out.println("\t");
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
                    text2 = text2.replace("NameRep", "Khuthadzo");
                }
                if (text2 != null){
                    text2 = text2.replace("InitialsRep", "KS");
                }
                if (text2 != null){
                    text2 = text2.replace("SurnameRep", "Mashapa");
                }
                if (text2 != null){
                    text2 = text2.replace("AddressLineRep", "831 Mamello street");
                }
                if (text2 != null){
                    text2 = text2.replace("CityRep", "Centurion");
                }

                if (text2 != null){
                    text2 = text2.replace("TitleRep", "MRS");
                }
                if (text2 != null){
                    text2 = text2.replace("CodeRep", "2000");
                }
                if (text2 != null){
                    text2 = text2.replace("Sequence of events", " ");
                }
                if (text2 != null) {
                    text2 = text2.replace("SignatureRep", "Khuthadzo Mashapa");
                }




                runs.get(j).setText(text2,0);
                System.out.println("    Run " + j + ": " + text2);
            }
            //start for name
            //start for address
            //start for address


        }
        //if(replaceCheck(text)){
//                List<XWPFRun> runs = paragraphs.get(i).getRuns();
//
//                for(int r = 0; r < runs.size(); r++){
//
//                    String runText = runs.get(r).getText(0);
//
//
//                    for(String keyWord : replacementFields){
//                        XWPFRun run = runs.get(r);
//                        runText = run.getText(0);
//                        if (runText == null){
//                            break;
//                        }
//                        if(runText.contains(keyWord)){
//
//                            String newText = null;
//                            if(keyWord == "AddressLineRep" || keyWord == "DistrictRep" ||
//                                    keyWord == "CityRep" || keyWord == "SurnameRep"){
//                                //System.out.println(run.getText(0));
//                                newText = runText.replace(keyWord, WordUtils.capitalize(details.get(keyWord).toLowerCase()));
//                                //System.out.println(newText);
//                            } else if(keyWord == "CaseNumberRep") {
//                                String year = transactions.get(0).get("IncidentDateRep").split("-")[0];
//                                String caseNumber = year+"-"+transactions.get(0).get(keyWord);
//
//                                newText = runText.replace(keyWord, caseNumber);
//                            } else if(keyWord == "IncidentDateRep"){
//                                SimpleDateFormat dt = new SimpleDateFormat("yyyy-MM-dd");
//
//                                Date date = null;
//                                try {
//                                    date = dt.parse(transactions.get(0).get(keyWord));
//                                } catch (ParseException e) {
//                                    e.printStackTrace();
//                                }
//
//                                SimpleDateFormat dt1 = new SimpleDateFormat("dd MMMMM YYYY");
//                                newText = runText.replace(keyWord, dt1.format(date));
//
//                            } else if(keyWord == "IncidentTimeRep" || keyWord == "MaskedCardNumberRep"){
//                                newText = runText.replace(keyWord, transactions.get(0).get(keyWord));
//                            } else if(keyWord == "ReasonRep") {
//                                newText = runText.replace(keyWord, caseLevel.get(transactions.get(0).get("CASE_LEVEL4")));
//                            } else if(keyWord == "RefundAmountRep"){
//                                newText = runText.replace(keyWord, "R"+ details.get( "RefundAmount").replace("R","").replace("r",""));
//                            } else if(keyWord == "OurRepName"){
//                                newText = runText.replace(keyWord, details.get( "OurRefName"));
//                            } else{
//                                //System.out.println(run.getText(0));
//                                System.out.println(keyWord);
//                                System.out.println(details.get(keyWord));
//                                newText = runText.replace(keyWord, details.get(keyWord));
//                            }
//                            run.setText(newText,0);
//                        }
//                    }
//
//
//
//                    // System.out.println("Run "+ r +": "+ runText);
//                }
//
//            }
//
//
//        }
//
//        List<XWPFTable> tables = document.getTables();
//
//        XWPFTable table = tables.get(0);
//        //table = newRow(table);
//
//        // this is the section that will have to deal with adding rows
//        // provided that more than 5 transactions are reported.
//
//        //table.createRow();
//
//
//        List<XWPFTableRow> tableRows = table.getRows();
//        //tableRows.remove(x);
//        //System.out.println("ROW SIZE: "+tableRows.size());
//
//        if (tableRows.size()-1 <= transactions.size()){
//            int insRow = transactions.size() - (tableRows.size()-1);
//            //System.out.println("INSERT ROW: "+insRow);
//            int i = 0;
//            while(i < insRow){
//                table.createRow();
//                i++;
//            }
//
//        }
//        for ( int r=0; r<tableRows.size();r++)
//        {
//            //System.out.println("Row "+ (r+1)+ ":");
//            XWPFTableRow tableRow = tableRows.get(r);
//            tableRow.setHeight(400);
//            List<XWPFTableCell> tableCells = tableRow.getTableCells();
//            for (int c=0; c<tableCells.size();c++)
//            {
//                //System.out.print("Column "+ (c+1)+ ": ");
//                XWPFTableCell tableCell = tableCells.get(c);
//                //tableCell.setText("TAE");
//
//                tableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                // this section is the basis of changing values in the table based on
//                // cell reference, easily modifiable.
//                if(r <= transactions.size() && r > 0){
//                    // System.out.println(c);
//                    switch(c){
//                        case 0:
//                            removeParagraphs(tableCell);
//                            tableCell.setText(transactions.get(r-1).get("TransDate"));
//                            break;
//                        case 1:
//                            removeParagraphs(tableCell);
//                            tableCell.setText(transactions.get(r-1).get("PostDate"));
//                            break;
//                        case 2:
//                            removeParagraphs(tableCell);
//                            if (!(transactions.get(r-1).get("TransTime").trim().equals("")) && (transactions.get(r-1).get("TransTime").trim().length()>8)) {
//                                tableCell.setText(transactions.get(r - 1).get("TransTime").replace(".", ":").substring(0, 8));
//                            }
//                                break;
//                        case 3:
//                            removeParagraphs(tableCell);
//                            tableCell.setText("R "+transactions.get(r-1).get("Amount"));
//                            break;
//                        case 4:
//                            removeParagraphs(tableCell);
//                            tableCell.setText(transactions.get(r-1).get("Location"));
//                            break;
//
//
//                    }
//                }
//
//                /*if(c==1 || r==1){
//                    removeParagraphs(tableCell);
//                    tableCell.setText("CHANGE");
//                }*/
//                String tableCellVal = tableCell.getText();
//
//                //System.out.println("tableCell.getText(" + (c) + "):" + tableCellVal);
//            }
//        }
//
        OutputStream out = null;
        try {
            out = new FileOutputStream(CompletedFileURL);
            document.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public boolean replaceCheck(String text){
        boolean flag = false;

        if (text.length() < 5){
            return false;
        }

        for (int i = 0; i < replacementFields.size(); i++){
            if(text.contains(replacementFields.get(i))){
                flag = true;
                //System.out.println(i);
                break;
            }
        }

        return flag;
    }

    public void removeParagraphs(XWPFTableCell tableCell) {
        int count = tableCell.getParagraphs().size();
        for(int i = 0; i < count; i++){
            tableCell.removeParagraph(i);
        }
    }

    public static void main(String[] args){
        letter test = new letter();
        test.replaceInfo("C:\\Users\\A239590\\Desktop\\CRC Letter New.docx");
    }



//    File file = new File("response.xml");
//        file.delete();
}
