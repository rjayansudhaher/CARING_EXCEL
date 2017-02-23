package com.excel.reading;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class Sortt {
	static ArrayList<String> list1=new ArrayList<String>();
	static Map<Integer, List> map = new HashMap<Integer, List>();
 	
	public List<Entry<Integer, List>> processOneSheet(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader( pkg );
		SharedStringsTable sst = r.getSharedStringsTable();

		XMLReader parser = fetchSheetParser(sst);

		InputStream sheet1 = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet1);
		parser.parse(sheetSource);
		
	//	String[] rowarray= new String[321];
		List<String> sublist = new ArrayList<String>();
		
		ArrayList<List<String>>  finallist = new ArrayList<List<String>>();
		
		
	    for(int k=1;k<=749;k++){
		/*for(int i=0; i<318;i++)
		{
			
			rowarray[i]=list.get((i+1)+(318*k));
			System.out.println("sa"+rowarray[i]);
			
		}*/
		
	    	for (int start = 316; start < list1.size(); start += 321) {
	            int end = Math.min(start + 321, list1.size());
	            finallist.add(list1.subList(start, end));
	           // finallist.add(sublist);	            
	           // System.out.println(sublist);
	            
	        }	
	    		
	    	map.put(k, finallist.get(k));
	    	
		
	//	System.out.println("Row no:"+k+""+finallist.get(k));
		
		
		
	    }
	    final int coltosort=252;
	    Set<Entry<Integer, List>> set = map.entrySet();
		List<Entry<Integer, List>> list = new ArrayList<Entry<Integer, List>>(set);
	    Collections.sort( list, new Comparator<Map.Entry<Integer, List>>()
	    {
	    	public int compare( Map.Entry<Integer, List> o1, Map.Entry<Integer, List> o2 )
	    	{return ((String) o1.getValue().get(coltosort)).compareTo( (String) o2.getValue().get(coltosort) );}
		} );
	    
	    for(Map.Entry<Integer, List> entry:list){
			System.out.println(entry.getKey()+" ==== "+entry.getValue().get(coltosort));}
	    
		
		sheet1.close();
		return list;
	}


	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser =
			XMLReaderFactory.createXMLReader(
					"org.apache.xerces.parsers.SAXParser"
			);
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	/** 
	 * See org.xml.sax.helpers.DefaultHandler javadocs 
	 */
	private static class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		
		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}
		
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {
			if(name.equals("c")) {
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
		
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			if(nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}
         
			 
			if(name.equals("v")) {
				list1.add(lastContents);
		
			}
		}

		public void characters(char[] ch, int start, int length)
				throws SAXException {
			lastContents += new String(ch, start, length);
		}
	}
	
	public static void main(String[] args) throws Exception {
		List<Entry<Integer, List>> list;
		XSSFWorkbook workbook = new XSSFWorkbook();
		Sortt example = new Sortt();
		list =example.processOneSheet("C:\\Users\\jayar29\\Desktop\\ACBS Lineage.xlsx");
		
		
		workbook= writeCollection(workbook, list);
		
		SXSSFWorkbook wb = new SXSSFWorkbook(workbook);
		   FileOutputStream out = new FileOutputStream( 
				      new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage_test.xlsx"));
		   				wb.write(out);
				      out.close();
		
	}

	private static XSSFWorkbook writeCollection(XSSFWorkbook workbook, List<Entry<Integer, List>> list) {
		
		XSSFSheet sheet = workbook.createSheet("Sorted");
		
		
		
		
		for (int index = 1,k=0; index <=list.size()+1 && k <list.size(); index++,k++) {
        	
        	Row r = sheet.createRow(index);
        	
        	Entry<Integer, List> rowobj=list.get(k);
        	
        	for(int j=0; j<rowobj.getValue().size();j++){
        		r.createCell(j).setCellValue(rowobj.getValue().get(j).toString());
        	
        	System.out.println(rowobj.getValue().get(j));
        	}
        	
        }
		
		
		return workbook;
	}
	
}
