package com.excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
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

import com.monitorjbl.xlsx.StreamingReader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ssso {
	
	 static int coltosort=0;

	public static void main(String[] args) throws IOException {
		
		  File file = new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage.xlsx");
		   FileInputStream fIP = new FileInputStream(file);
		   XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		   
		   /*InputStream is = new FileInputStream(new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage.xlsx"));
		   XSSFWorkbook  workbook = (XSSFWorkbook) StreamingReader.builder()
		            .rowCacheSize(100)    
		            .bufferSize(4096)     
		            .open(is); */ 
		   
		   
		   List<Entry<Integer, String[]>> list;
		   
		   list= colsort(workbook);
		   
		   
		   workbook= writeCollection(workbook,list);
		   
//		   for(Map.Entry<Integer, String[]> entry:list){
//				System.out.println(entry.getKey()+" ==== "+entry.getValue()[coltosort]);}
		   
		   
		   SXSSFWorkbook wb = new SXSSFWorkbook(workbook);
		   FileOutputStream out = new FileOutputStream( 
				      new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage_test.xlsx"));
		   				wb.write(out);
				      out.close();
		   
		   
		
	}
	
	

	private static List<Entry<Integer, String[]>> colsort(XSSFWorkbook workbook) {
		
		 int sortingcol=0;
		 Map<Integer, String[]> map = new HashMap<Integer, String[]>();
	        
		 XSSFSheet sheet = workbook.getSheetAt(0);
		 int lastRow = sheet.getLastRowNum();
        
		 
		Row row=sheet.getRow(1);
		Cell cell=row.getCell(272);
		
		if(!cell.toString().equals("N/A"))
		{
			sortingcol=272;
		}else{
			sortingcol=252;
		}
		
	   coltosort=sortingcol;
		
		for (int j=1; j<=lastRow;j++)
		{
			
			String[] cellsinRow = null;
	    	int i=0;
    	
	    	int cellCount = sheet.getRow(j).getLastCellNum();
            System.out.println("CellCount"+cellCount);
	    	
	    	cellsinRow	=	new String[cellCount];

	    	Row DataRow=sheet.getRow(j);
	    	

	    	Iterator<Cell> HeaderCellIterator1 = DataRow.cellIterator();
	    	
	    	while(HeaderCellIterator1.hasNext())
	    	{
	    		Cell CurrentRowCells1 = HeaderCellIterator1.next();
	    		
	    		
	    		switch( CurrentRowCells1.getCellType())
		    	{
	    		
	    			case Cell.CELL_TYPE_NUMERIC:
	    			
		    			cellsinRow[i] = CurrentRowCells1.toString().replaceAll("\\.?0*$", "");   //Removing Trailing Zeros and assigning to the String
		    			break;
	    			
		    		case Cell.CELL_TYPE_BLANK:
		    			
		    			cellsinRow[i] = "";
			    	
		    		case Cell.CELL_TYPE_STRING:
		    			
		    			cellsinRow[i] = CurrentRowCells1.toString();
			    		break;
		    		
		    		case Cell.CELL_TYPE_BOOLEAN:
		    			
		    			cellsinRow[i] = CurrentRowCells1.toString();
			    		break;
			    		
		    		case Cell.CELL_TYPE_ERROR:
		    			
		    			cellsinRow[i] = CurrentRowCells1.toString();
			    		break;
		    		
		    		case Cell.CELL_TYPE_FORMULA:
		    			
		    			/*cellsinRow[i] = CurrentRowCells1.toString();
			    		break;*/
		    			
		    			 switch(CurrentRowCells1.getCachedFormulaResultType()) {
                         case Cell.CELL_TYPE_NUMERIC:
                             System.out.println("Last evaluated as: " + CurrentRowCells1.getNumericCellValue());
                             cellsinRow[i] = CurrentRowCells1.toString();
                                  break;
                             
                         case Cell.CELL_TYPE_STRING:
                             System.out.println("Last evaluated as \"" + CurrentRowCells1.getRichStringCellValue() + "\"");
                             break;
		    			 }
		    		
			    	default:
			    		System.out.println("Defaul");
			    		cellsinRow[i]	=	"";					    	
		    	}
	    		
	    		i++;
	    	}
		
			
	    	map.put(j, cellsinRow);
						
	
		}
	    
       
		Set<Entry<Integer, String[]>> set = map.entrySet();
		List<Entry<Integer, String[]>> list = new ArrayList<Entry<Integer, String[]>>(set);
	    Collections.sort( list, new Comparator<Map.Entry<Integer, String[]>>()
	    {
	    	public int compare( Map.Entry<Integer, String[]> o1, Map.Entry<Integer, String[]> o2 )
	    	{return (o1.getValue()[coltosort]).compareTo( o2.getValue()[coltosort] );}
		} );
	    
	    return list;
	    
		/*for(Map.Entry<Integer, String[]> entry:list){
		System.out.println(entry.getKey()+" ==== "+entry.getValue()[coltosort]);}*/

		
		
	}
	
	
	private static XSSFWorkbook writeCollection(XSSFWorkbook workBook,
			List<Entry<Integer, String[]>> list) {
		
		  try {
		        //FileInputStream fileInputStream = new FileInputStream(file);
		       // XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
		        XSSFSheet sheet = workBook.getSheetAt(0);
		        int rows = sheet.getLastRowNum();
		        System.out.println("Coming here");
		        
		        /*for(Map.Entry<Integer, String[]> entry:list){
				System.out.println(entry.getKey()+" ==== "+entry.getValue()[coltosort]);}*/
		        
		        
		        /*for(Map.Entry<Integer, String[]> entry:list){
		        	   int i=1;
		        	   Row r = sheet.getRow(i);
		        	  
		        	   for(int j=1; j<r.getLastCellNum();j++){
		        	   r.getCell(j).setCellValue(entry.getValue()[j]);
		        	   System.out.println(i+".."+entry.getValue()[j]);
		        	   }
		        	}*/
		      
		        
		        for (int index = 1,k=0; index <=rows && k <list.size(); index++,k++) {
		        	
		        	Row r = sheet.getRow(index);
		        	
		        	Entry<Integer, String[]> rowobj=list.get(k);
		        	
		        	for(int j=0; j<rowobj.getValue().length;j++){
		        		r.getCell(j).setCellValue(rowobj.getValue()[j]);
		        	
		        	System.out.println(rowobj.getValue()[j]);
		        	}
		        	
		        }
		        
		
		    } catch (Exception e) {
		       e.printStackTrace();
		    }
		return workBook;
		
		
	}
}
