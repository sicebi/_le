package com.company;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.formula.functions.Replace;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
//import org.apache.commons.text.WordUtils;
import java.awt.*;
import java.awt.font.TextAttribute;
import java.io.IOException;
import java.io.Reader;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.AttributedString;
import java.time.LocalDate;
import java.util.*;

import org.apache.poi.xwpf.usermodel.*;
import sun.security.util.AuthResources_de;

import javax.xml.bind.annotation.XmlType;
import java.io.*;
import java.util.List;
import java.util.regex.Pattern;

import static javafx.application.Platform.exit;



class decodeHtml {
    List<String> replacementFields;
    Map<String, String> caseLevel;

    public void decode(String ReferenceRep,String DateRep, String TotalfraudRep, String TotalrefundRep, String TotallossRep, List<String> seqEvents) {


        replacementFields = new ArrayList<String>();
        replacementFields.add("SeqEventsRep");

        XWPFNumbering numbering = null;
        XWPFNum num = null;
        BigInteger numID = null;
        int numberingID = -1;
        String CompletedFileURL = "C:\\Users\\\\A233553\\Desktop\\Completed latter\\" + ReferenceRep + ".docx";
        InputStream fis = null;
        try {
            fis = new FileInputStream("C:\\Users\\A233553\\Desktop\\src\\CRC Letter New.docx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XWPFDocument document = null;
        try {
            document = new XWPFDocument(fis);
        }  catch (IOException ioException) {
            throw new RuntimeException("One Opening The Letter",ioException);
        }
//                getting information from the table_________________________________________________________
        //                getting information from the table
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            for (XWPFTableRow row : table.getRows()) {

                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph tableP : cell.getParagraphs())
                        for (XWPFRun run : tableP.getRuns()) {
                            String text = run.getText(0);
                            if (text.contains("DateRep")) {
                                text = text.replace("DateRep", DateRep);
                            }
                            if (text.contains("ReferenceRep")) {
                                text = text.replace("ReferenceRep", ReferenceRep);
                            }
                            if (text.contains("TotalfraudRep")) {
                                text = text.replace("TotalfraudRep", TotalfraudRep);
                            }
                            if (text.contains("TotalrefundRep")) {
                                text = text.replace("TotalrefundRep", TotalrefundRep);
                            }
                            if (text.contains("TotallossRep")) {
                                text = text.replace("TotallossRep", TotallossRep);
                            }
                            run.setText(text, 0);
//                            System.out.println("cell text------ ++++++ " + text);

                        }
                }

                System.out.println(" ");
            }
        }

        ///_____________________________________________________________________________________________________________
        numbering = document.getNumbering();
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (int i = 0; i < paragraphs.size(); i++){
            String text = paragraphs.get(i).getText();
            System.out.println("Paragraph "+ i +": "+ text);
            Boolean delFlag = text.contains("SeqEventsRep");
            Boolean ifEmpty = false;
            int count = 0;
            if(replaceCheck(text) || i == 24){
                List<XWPFRun> runs = paragraphs.get(i).getRuns();

                for(int r = 0; r < runs.size(); r++){

                    String runText = runs.get(r).getText(0);
                    System.out.println("###########Run " + r + ": " + runText);


                    for(String keyWord : replacementFields){
                        XWPFRun run = runs.get(r);
                        runText = run.getText(0);
                        if (runText == null){
                            break;
                        }
                        if(runText.contains(keyWord)) {

                            String newText = "";
//                            if (keyWord == "AddressLineRep" || keyWord == "DistrictRep" ||
//                                    keyWord == "CityRep" || keyWord == "SurnameRep") {
//                                //System.out.println(run.getText(0));
//                                newText = runText.replace(keyWord, WordUtils.capitalize(details.get(keyWord).toLowerCase().trim()));
//                                //System.out.println(newText);
//                            } else
                            if (keyWord == "RefNumberRep") {
                                String refNumber = ReferenceRep;

                                newText = runText.replace(keyWord, refNumber);
                            } else if (keyWord == "SeqEventsRep"){

                                numID = paragraphs.get(i).getNumID();
                                String font = run.getFontFamily();
                                int fontSize = run.getFontSize();
                                count = 0;
                                String type = seqEvents.get(seqEvents.size()-2);
                                char bulletChar = seqEvents.get(seqEvents.size()-1).charAt(0);
                                if(type.equals("Empty")){
                                    ifEmpty = true;

                                } else if (type.equals("Partial")){

                                    for (int l = seqEvents.size() - 3; l >= 0; l--) {
                                        count++;
                                        char firstChar = seqEvents.get(l).charAt(0);
                                        XmlCursor cursor = paragraphs.get(i).getCTP().newCursor();
                                        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
                                        if(firstChar == bulletChar) {
                                            //System.out.println("FOUND ONE HERE _____________");
                                            paragraph.setNumID(numID);
                                        }
                                        XWPFRun runInner = paragraph.createRun();
                                        runInner.setBold(false);
                                        runInner.setFontSize(fontSize);
                                        runInner.setFontFamily(font);
                                        if(firstChar == bulletChar) {

                                            String deBulleted = seqEvents.get(l).substring(1).trim();
//                                            System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&1" + "\u2023   " + deBulleted );
                                            runInner.setText("\u2023   " + deBulleted);

                                        } else{

//                                            System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&2" + seqEvents.get(l) );
                                            runInner.setText("\u2023   " + seqEvents.get(l));

                                        }

                                    }

                                } else {
                                    for (int l = seqEvents.size() - 3; l >= 0; l--) {
                                        count++;
                                        XmlCursor cursor = paragraphs.get(i).getCTP().newCursor();
                                        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
                                        paragraph.setNumID(numID);
                                        XWPFRun runInner = paragraph.createRun();
                                        runInner.setBold(false);
                                        runInner.setFontSize(fontSize);
                                        runInner.setFontFamily(font);
//                                        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&3" + seqEvents.get(l) );
                                        runInner.setText("\u2023   " + seqEvents.get(l));
                                    }
                                }
//                                newText = runText.replace(keyWord, "I NEED TO FIGURE OUT BULLET LISTS");
                            } else{
                                //System.out.println(run.getText(0));
                                //System.out.println(keyWord);
                                // System.out.println(details.get(keyWord));
//                                newText = runText.replace(keyWord, details.get(keyWord));
                            }
//                            System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&4" + newText );
                            run.setText(newText,0);
                        }
                    }



                    // System.out.println("Run "+ r +": "+ runText);
                }

            }
            if(delFlag){
//                int pPos = document.getPosOfParagraph(paragraphs.get(i+count));
//                document.removeBodyElement(pPos);
                exit();
            }
            if(ifEmpty){
//                 System.out.println("EMPTY ONE OVER HERE <----------------------");
                 exit();
            }

        }
        //______________________________________________________________________________________________________________
//        getting the paragraphs only

        OutputStream out = null;
        try {
            out = new FileOutputStream(CompletedFileURL);
            document.write(out);
            out.close();
        } catch (IOException ioException) {
            throw new RuntimeException("The writting the Letter ",ioException);
        }
    }

    public boolean replaceCheck(String text) {
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


    public static List<String> getSequenceEvents2(String Html){

        List<String> seqEvents = new ArrayList<String>();
        String bodyText;
        String childText;
        Document doc = Jsoup.parse(Html, "UTF-8");

        Element body = doc.getElementsByTag("body").first();
        bodyText = body.text();
        Element wrapCheck = body.children().first();
//        System.out.println("wrapCheck = ====================" + wrapCheck);
        if (wrapCheck == null && bodyText != null){
            seqEvents.add(body.text().replace(";", ""));

            //            return seqEvents;
        }else{
        childText = wrapCheck.text();
            while (bodyText.equals(childText) && wrapCheck.childNodeSize()>0){
                body = wrapCheck;
                wrapCheck = wrapCheck.children().first();
                childText = wrapCheck.text();
            }
//            System.out.println("wrapCheck.text() +++++++++++++++++++++++++++++++++++++ " + wrapCheck.text());
        }


        Elements events = body.children();

        for (Element event : events) {

            if(event.getElementsByTag("div").size()>1){
                for (Element child: event.children()){
                    if(child.text().length() > 0) {
                        seqEvents.add(child.text());
                    }
                }

            } else if (event.text().length() > 0) {
                seqEvents.add(event.text());
            }


        }
        if(seqEvents.size() != 0 ) {
            Map<Character, Map<String, List<Integer>>> firstChars = new HashMap<Character, Map<String, List<Integer>>>();
            for (int i = 0; i < seqEvents.size(); i++) {
                char firstChar = seqEvents.get(i).charAt(0);

                if(firstChars.containsKey(firstChar)){
                    Map<String, List<Integer>> thisInfo = firstChars.get(firstChar);
                    int count = thisInfo.get("count").get(0) + 1;
                    thisInfo.get("count").set(0, count);
                    thisInfo.get("pos").add(i);
                } else {
                    List<Integer> pos = new ArrayList<Integer>();
                    pos.add(i);
                    List<Integer> count = new ArrayList<Integer>();
                    count.add(1);

                    Map<String, List<Integer>> newInfo = new HashMap<String, List<Integer>>();
                    newInfo.put("count", count);
                    newInfo.put("pos", pos);
                    firstChars.put(firstChar, newInfo);
                }

            }
            Boolean foundFlag = false;
            for(char key : firstChars.keySet()){
                String firstChar = Character.toString(key);
                if(!Pattern.matches("[a-zA-Z]", firstChar)){
                    int count = firstChars.get(key).get("count").get(0);

                    if(seqEvents.size() - count < 5){
                        seqEvents.add("Partial");
                        seqEvents.add(firstChar);
                        foundFlag = true;
                        break;
                    }
                }
            }

            if(!foundFlag){
                seqEvents.add("None");
                seqEvents.add("Filler");
            }

        } else {
            seqEvents.add("Empty");
            seqEvents.add("Filler");
        }
//        System.out.println("seqEvents ---------------------------- " + seqEvents);
        return seqEvents;
    }
    public static class Details{
        private String _ReferenceRep;
        private String _StatusRep;
        private String _TotalfraudRep;
        private String _TotalrefundRep;
        private String _TotallossRep;
        private String _GISstatusRep;
        private String _TS_BPID_Rep;
        private String _EventRep;
        private String _DateRep;



        public void set_ReferenceRep(String _ReferenceRep) {
            this._ReferenceRep = _ReferenceRep;
        }

        public void set_StatusRep(String _StatusRep) {
            this._StatusRep = _StatusRep;
        }

        public void set_TotalfraudRep(String _TotalfraudRep) {
            this._TotalfraudRep = _TotalfraudRep;
        }

        public void set_TotalrefundRep(String _TotalrefundRep) {
            this._TotalrefundRep = _TotalrefundRep;
        }

        public void set_TotallossRep(String _TotallossRep) {
            this._TotallossRep = _TotallossRep;
        }

        public void set_GISstatusRep(String _GISstatusRep) {
            this._GISstatusRep = _GISstatusRep;
        }

        public void set_TS_BPID_Rep(String _TS_BPID_Rep) {
            this._TS_BPID_Rep = _TS_BPID_Rep;
        }

        public void set_EventRep(String _EventRep) {
            this._EventRep = _EventRep;
        }
        public String get_ReferenceRep() {
            return _ReferenceRep;
        }

        public String get_StatusRep() {
            return _StatusRep;
        }

        public String get_TotalfraudRep() {
            return _TotalfraudRep;
        }

        public String get_TotalrefundRep() {
            return _TotalrefundRep;
        }

        public String get_TotallossRep() {
            return _TotallossRep;
        }

        public String get_GISstatusRep() {
            return _GISstatusRep;
        }

        public String get_TS_BPID_Rep() {
            return _TS_BPID_Rep;
        }

        public String get_EventRep() {
            return _EventRep;

        }


        public String get_DateRep() {
            return _DateRep;
        }

        public void set_DateRep(String _DateRep) {
            this._DateRep = _DateRep;
        }
    }

    public static class CsvReader{
        Details _Details = new Details();
        final String SAMPLE_CSV_FILE_PATH = "C:\\Users\\A233553\\Desktop\\src\\letters.csv";
        {
            try {
                Reader reader = Files.newBufferedReader(Paths.get(SAMPLE_CSV_FILE_PATH));
                CSVReader csvReader = new CSVReaderBuilder(reader).withSkipLines(1).build();
                String[] nextRecord;

                while ((nextRecord = csvReader.readNext()) != null) {
                    _Details.set_ReferenceRep(nextRecord[0]);
                    _Details.set_StatusRep(nextRecord[1]);
                    _Details.set_TotalfraudRep(nextRecord[2]);
                    _Details.set_TotalrefundRep(nextRecord[3]);
                    _Details.set_TotallossRep(nextRecord[4]);
                    _Details.set_GISstatusRep(nextRecord[5]);
                    _Details.set_TS_BPID_Rep(nextRecord[6]);
                    _Details.set_EventRep(nextRecord[7]);
                    LocalDate today = LocalDate.now();
                    _Details.set_DateRep(String.valueOf((today)));

                    System.out.println("========================================================");
                    System.out.println("Reference : " + _Details.get_ReferenceRep());
                    System.out.println("Date : " + _Details.get_DateRep());
                    System.out.println("Amount R : " + _Details.get_TotalfraudRep());
                    System.out.println("Recovered R : " +_Details.get_TotalrefundRep());
                    System.out.println("Customer Loss R : " +_Details.get_TotallossRep());
                    System.out.println("Event  : " + _Details.get_EventRep());
                    System.out.println("========================================================");

                decodeHtml _decodeHtml = new decodeHtml();
                _decodeHtml.decode(_Details.get_ReferenceRep(),_Details.get_DateRep(),_Details.get_TotalfraudRep(), _Details._TotalrefundRep,_Details._TotalrefundRep, getSequenceEvents2(_Details.get_EventRep()));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public static void main(String[] args) {

        CsvReader csv = new CsvReader();

    }


}
