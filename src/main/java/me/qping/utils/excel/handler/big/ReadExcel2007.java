package me.qping.utils.excel.handler.big;

import me.qping.utils.excel.common.RowConsumer;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReadExcel2007 extends DefaultHandler {

    public void processOneSheet(InputStream inputStream, int sheetNo, int dataRowNum, RowConsumer rowConsumer) throws Exception {
        OPCPackage pkg = OPCPackage.open(inputStream);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        ContentHandler handler = new SheetHandler(sst, rowConsumer, dataRowNum);
        parser.setContentHandler(handler);

//        XMLReader parser = fetchSheetParser(sst, rowConsumer, dataRowNum);

        // To look up the Sheet Name / Sheet Order / rID,
        //  you need to process the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet1 = r.getSheet("rId" + (sheetNo + 1));
        InputSource sheetSource = new InputSource(sheet1);
        parser.parse(sheetSource);
        sheet1.close();
    }



    public static int AAtoNumber(String str) {
        int result = 0;
        int len = str.length();
        long base = 1;
        long N = 26;
        for (int i = len - 1; i >= 0; i--) {
            int index = str.charAt(i) - 64;
            result += index * base;
            base *= N;
        }
        return result;
    }

    private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private RowConsumer rowConsumer;
        private int dataRowNum;


        private SheetHandler(SharedStringsTable sst, RowConsumer rowConsumer, int dataRowNum) {
            this.rowConsumer = rowConsumer;
            this.sst = sst;
            this.dataRowNum = dataRowNum;
        }

        Map<Integer,Object> rowData = new HashMap<>();
        long row = -1;
        int col = -1;
        Pattern pattern = Pattern.compile("^([a-zA-Z]+)([0-9]+$)");


        public void setRowCol(String label){

            Matcher mather = pattern.matcher(label);
            long rowCurrent;
            int colCurrent;

            if(mather.matches()){
                try{
                    rowCurrent = Long.parseLong(mather.group(2)) - 1;
                    colCurrent = AAtoNumber(mather.group(1)) - 1;
                }catch (Exception ex){
                    throw new RuntimeException("cell reference is illegal:" + label);
                }
            }else{
                throw new RuntimeException("cell reference is illegal:" + label);
            }

            if(row != rowCurrent){

                if(row != -1 && row >= dataRowNum){
                    rowConsumer.execute(rowData, row , col);
                }

                rowData = new HashMap<>();
                row = rowCurrent;
            }

            col =  colCurrent;
        }


        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            // c => cell
            if(name.equals("c")) {
                // Print the cell reference
                setRowCol(attributes.getValue("r"));
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                if(cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name) throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v")) {
                rowData.put(col, lastContents);
            }
        }

        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }

        public void endDocument () throws SAXException {
            if(row >= dataRowNum){
                rowConsumer.execute(rowData, row , col);
            }
        }
    }


}
