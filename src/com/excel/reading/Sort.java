package com.excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sort {

	public static void main(String[] args) throws IOException {
		
	
		// Need to change the path accordingly
		File file = new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage.xlsx");
		   FileInputStream fIP = new FileInputStream(file);
		   XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		   
		    
		   workbook=sortCells(workbook);
		   SXSSFWorkbook wb = new SXSSFWorkbook(workbook);
		   FileOutputStream out = new FileOutputStream( 
				      new File("C:\\Users\\jayar29\\Desktop\\ACBS Lineage_test.xlsx"));
		   				wb.write(out);
				      out.close();

	}
	
	private static XSSFWorkbook sortCells(XSSFWorkbook workbook) {
		
		 int sheetCount = workbook.getNumberOfSheets();
		 for(int i=sheetCount-1;i>=0 ;i--){
			 
			 int sortingcol=0;
	        
			 XSSFSheet sheet = workbook.getSheetAt(i);
			 int lastRow = sheet.getLastRowNum();
	        
			 
			Row row=sheet.getRow(1);
			Cell cell=row.getCell(272);
			
			if(!cell.toString().equals("N/A"))
			{
				sortingcol=272;
			}else{
				sortingcol=252;
			}
			
	        
	        Row row1=null;
	        Row row2=null;
	        
	        boolean sorting = true;
	        
	        /*while (sorting == true) {
	            sorting = false;
	            for (Row row : sheet) {
	                // skip if this row is before first to sort
	                if (row.getRowNum()<0) continue;
	                // end if this is last row
	                if (lastRow==row.getRowNum()) break;
	                Row row2 = sheet.getRow(row.getRowNum()+1);
	                if (row2 == null) continue;
	                String firstValue = (row.getCell(6) != null) ? row.getCell(6).getStringCellValue() : "";
	                String secondValue = (row2.getCell(6) != null) ? row2.getCell(6).getStringCellValue() : "";
	                //compare cell from current row and next row - and switch if secondValue should be before first
	                
	                System.out.println("@@"+secondValue);
	                System.out.println("##"+firstValue);
	                if (secondValue.compareToIgnoreCase(firstValue)<0) {  
	                	
	                	System.out.println("sooooooort");
	                    sheet.shiftRows(row2.getRowNum(), row2.getRowNum(), -1);
	                    sheet.shiftRows(row.getRowNum(), row.getRowNum(), 1);
	                    sorting = true;
	                }
	            }
	        }
	        */
	        
	        try{
	        	for(int j=1; j<=sheet.getLastRowNum();j++){
	        		
	        	row1=sheet.getRow(j);
	        	Cell ce1=row1.getCell(sortingcol);
	        	
	        	for (int k=j+1;k<=sheet.getLastRowNum();k++)
	        	{
	        		row2=sheet.getRow(k);
	        		Cell ce2=row2.getCell(sortingcol);
	        		
	        		System.out.println("@@"+ce1.toString());
	                System.out.println("##"+ce2.toString());
	        		if(ce2.toString().compareToIgnoreCase(ce1.toString())<0)
	        		{
	        			System.out.println("Soort");
	        			for(int l=0;l<row1.getLastCellNum();l++)
	        			{
	        				if(row1.getCell(l) !=null && row2.getCell(l) !=null){
	        				String s=row1.getCell(l).toString();
	        				row1.getCell(l).setCellValue(row2.getCell(l).toString());
	        				row2.getCell(l).setCellValue(s);
	        				}
	        				
	        			}
	        				        			
	        		}
	        	}
	        	
	        	
			}
	        }catch(Exception e)
	        {
	        	e.printStackTrace();
	        	continue;
	        }
	        
	        
		 }
		
		return workbook;
	}

}
