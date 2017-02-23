/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author PR355016
 */
public class ExcelMergerNew {
	
	public static int count=0;
    public static String[] totalworkbooks=new String[1000];
    public static int bookcount=0;
    public static int sheetcount=0;
    public static int finalrowcount=0;
    public static XSSFWorkbook workbookfinal = null;
    
	
    
    public static XSSFWorkbook mergeExcel(XSSFWorkbook workbook, int cellCounter, String[] sheetname){
        
        
        if(workbook.getNumberOfSheets()<=1){
             return workbook;
        }else{
            workbook = cellCompare(workbook);
           // workbook = sortCells(workbook);
          //  writeExcel(workbook, "ASD");
            
            workbook = rowCompare(workbook);
                      
            XSSFWorkbook listWorkbook = createExcel(workbook, cellCounter, sheetname);
          //  writeExcel(listWorkbook, "nnnn");
            return listWorkbook;
           
        }
        
    }
    


public static XSSFWorkbook cellCompare(XSSFWorkbook workbook){
        
        for(int i=0;i<workbook.getNumberOfSheets();i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            
            for(int j=1;j<sheet.getLastRowNum();j++){
                 Row row1 = sheet.getRow(j);
                 for(int k=j+1;k<=sheet.getLastRowNum();k++){
                 Row row2 = sheet.getRow(k);
                 if(row1 !=null && row2 != null){
                 if(row1.getCell(5).toString().toLowerCase().matches("exp_.*") || row1.getCell(6)==null){
                     XSSFCell expValue = (XSSFCell) row1.getCell(5);
                     expValue.setCellValue("N/A");
                     
                     XSSFCell expValue1 = (XSSFCell) row1.getCell(6);
                     expValue1.setCellValue("N/A");
                 }
                 
                 if(row2.getCell(5).toString().toLowerCase().matches("exp_.*")|| row2.getCell(6)==null){
                     XSSFCell expValue = (XSSFCell) row2.getCell(5);
                     expValue.setCellValue("N/A");
                     XSSFCell expValue1 = (XSSFCell) row2.getCell(6);
                     expValue1.setCellValue("N/A");
                 }
                 }
                 
                 try{
                if( row1.getCell(12).toString().equalsIgnoreCase(row2.getCell(12).toString()) && row1.getCell(6).toString().equalsIgnoreCase(row2.getCell(6).toString())){
                    
                    Cell calcCell1 = row1.getCell(20);
                    Cell calcCell2 = row2.getCell(20);
                    Cell targCell1 = row1.getCell(11);
                    Cell targCell2 = row2.getCell(11);
                    String calc = calcCell1+","+calcCell2;
                    String targ = targCell1+","+targCell2;
                    calcCell1.setCellValue(calc);
                    targCell1.setCellValue(targ);
                    
                   //System.out.println(calc);
                    
                    sheet.removeRow(row2);
                    sheet.shiftRows(k+1, sheet.getLastRowNum(), -1);
                    
                    j--;
                    //k--;
                }
                
                
                }catch(Exception e){
                    
                }
                
                 }
                 
            }
            
        }
        
        return workbook;
    }

private static XSSFWorkbook sortCells(XSSFWorkbook workbook) {
	
	 int sheetCount = workbook.getNumberOfSheets();
	 for(int i=sheetCount-1;i>=0 ;i--){
        
        XSSFSheet sheet = workbook.getSheetAt(i);
        Row row1=null;
        Row row2=null;
        
        boolean sorting = true;
        int lastRow = sheet.getLastRowNum();
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
        	Cell ce1=row1.getCell(72);
        	
        	for (int k=j+1;k<=sheet.getLastRowNum();k++)
        	{
        		row2=sheet.getRow(k);
        		Cell ce2=row2.getCell(72);
        		
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
        	//e.printStackTrace();
        	continue;
        }
        
        
	 }
	
	return workbook;
}
    
    public static XSSFWorkbook rowCompare(XSSFWorkbook workbook){
        
        for(int i=workbook.getNumberOfSheets()-1;i>0;i--){
            XSSFSheet sheet1 = workbook.getSheetAt(i);
            XSSFSheet sheet2 = workbook.getSheetAt(i-1);
            //System.out.println("Haii");
           // if(sheet1.getLastRowNum() == sheet2.getLastRowNum()){
            
           int sheeet1lastrow=sheet1.getLastRowNum();
            int sheeet2lastrow=sheet2.getLastRowNum();
            
            
            if(sheetcount==0 || sheetcount<sheeet1lastrow)
            		{sheetcount=sheeet1lastrow;}
            else
            {sheetcount=sheeet2lastrow;}
            
            HashMap<String, String[]> tgtmap=new HashMap<>();
            
            
            for(int s=sheet2.getLastRowNum()+1;s<sheet1.getLastRowNum();s++)
            {
          	  if(sheet2.getRow(s) !=null)
          	  {
          		  
          	  }else
          	  {
          		 Row r1= sheet2.createRow(s);
          		  
          		  for(int t=0;t<26;t++)
          		  {
          			  r1.createCell(t).setCellValue("N/A");
          		  }
          		  
          	  }
            }
            
                        
            
                for(int j=1;j<=sheeet1lastrow ;j++){
                	
                	System.out.println("roooooooooo"+sheeet1lastrow);
                if(sheet1.getRow(j) !=null && sheet2.getRow(j) !=null){
                	
                	String[] strlist = new String[29];
                	
                    Row row1 = sheet1.getRow(j);
                    Row row2=  sheet2.getRow(j);
                    System.out.println("zzzzzz"+row1.getCell(6));
                    System.out.println("zzzzzz1"+row2.getCell(12));
                    
                    if(row1.getCell(6) !=null && row2.getCell(12) !=null ){
                        Cell cell1 = row1.getCell(6);
                        Cell cell2 = row2.getCell(12);
                        
                        
                        for(int a=0; a<row2.getLastCellNum();a++)
                        {
                        	try{
                        	strlist[a]=row2.getCell(a).toString();	
                        	}catch(Exception e){continue;}
                        }
                       
                       System.out.println("Adding"+cell1.toString());
                       tgtmap.put(cell2.toString(),strlist);
                       
                    
                    
                    try{
                    
                    if(row1.getCell(6).toString().equals("N/A"))
                    {
                    	
                    //	sheet2.removeRow(row2);                        	
                        //	sheet1.shiftRows(j+1, sheet1.getLastRowNum(), -1);
                        //	--j;
                    	int sht1lastrow=sheeet2lastrow;
                    	Row newrow= sheet2.createRow(sht1lastrow);
                    	
                    	for(int d=0; d<row2.getLastCellNum(); d++)
                    	{
                    		newrow.createCell(d).setCellValue("N/A");
                    		
                    		Cell valuecell= row2.getCell(d);
                    		Cell dummycell= newrow.getCell(d);
                    		
                    		String save= dummycell.toString();
                    		dummycell.setCellValue(valuecell.toString());
                    		valuecell.setCellValue(save);
                    		
                    		
                    	}
                    	
                    	
                    	/*
                    	for(int a=0;a<row2.getLastCellNum();a++)
                    	
                    			row2.getCell(a).setCellValue("N/A");*/
                    	//row2.getCell(12).setCellValue("N/A");
                    	
                        	//continue;
                   
                    	
                    }
                    
                    }catch(Exception e){e.printStackTrace();}
                 //   {continue;}
                    
                   
                  
                   
                   
                   
                    
                    boolean checked=false;
                    
                   // System.out.println("^^^^^^^"+cell1.toString());
                    
                    try{
                    if(!(cell1.toString().equalsIgnoreCase(cell2.toString()))){
                    	
                                        	                    	
                      //  System.out.print(cell1 + "----"+j+"----");
                      //     System.out.println("####"+cell2);
                         for(int k = j+1;k<=sheeet2lastrow;k++){
                            
                               Row dummyRow = sheet2.getRow(k);
                               Cell dummyCell = dummyRow.getCell(12);
                               
                               if(cell1.toString().equalsIgnoreCase(dummyCell.toString())){

                                  // System.out.print(cell1 + "----"+k+"----");
                                 //  System.out.println("&&&&&&"+dummyCell);
                            	  
                            	   checked=true;
                            	   
                            	   String[] asd=new String[dummyRow.getLastCellNum()];
                            	   
                            	   for(int f=0; f<dummyRow.getLastCellNum();f++)
                            	   {
                            		 asd[f]= dummyRow.getCell(f).toString() ;  
                            	   }
                            	   
                            	   tgtmap.put(dummyCell.toString(),asd);
                                   
                                   for(int l=0;l<dummyRow.getLastCellNum();l++){
                                   try{   
                                	   
                                	    
                                	   Cell cellOne = row2.getCell(l);
                                       Cell cellTwo = dummyRow.getCell(l);
                                       String dummyStr = cellTwo.toString();
                                     //  .out.println("cellOne :: "+cellOne+" cellTwo :: "+cellTwo+" dummyCell :: "+dummyStr);
                                       
                                       if(cellOne !=null){
                                       cellTwo.setCellValue(cellOne.toString());
                                       cellOne.setCellValue(dummyStr);}
                                       }catch(Exception e){
                                    	   e.printStackTrace();
                                       continue;
                                   }
                                   }
                                   
                               }
                               
                               
                        }
                         
                         
                        if(!checked)
                        {
                        	
                        	
                        	++count;
                        	
                        	
                        	if(tgtmap.containsKey(row1.getCell(6).toString()))
                        	{
                        		System.out.println("**->"+row1.getCell(6).toString() + "-->"+row2.getCell(12).toString());
                        		Row endrow= sheet2.createRow(sheeet2lastrow);
                        		Row Swapingrow= sheet2.getRow(j);
                        		String[] strlist1=new String[Swapingrow.getLastCellNum()];
                        		strlist1=tgtmap.get(row1.getCell(6).toString());
                        		
                        		System.out.println("Deleting"+row1.getCell(6).toString());
                        		for(int s=0;s<strlist1.length;s++)
                        		{
                        			System.out.println(strlist1[6]);
                        		//	System.out.println("******************"+strlist1[s]);
                        		}
                        		
                        		
                        		for(int l=0;l<26;l++){
                                    try{    Cell cellOne = Swapingrow.getCell(l);
                                    
                                    	endrow.createCell(l);
                                    	endrow.getCell(l).setCellValue("N/A");
                                    	Cell cellTwo = endrow.getCell(l);
                                       // String dummyStr = cellTwo.toString();
                                        
                                       // System.out.println("*****"+cellOne.toString());
                                        //System.out.println(dummyStr);
                                        
                                        
                                        
                                        //System.out.println("cellOne :: "+cellOne+" cellTwo :: "+cellTwo+" dummyCell :: "+dummyStr);
                                        if(cellOne != null)
                                    	cellTwo.setCellValue(cellOne.toString());
                                        cellOne.setCellValue(strlist1[l]);
                                        
                                     //  System.out.println(cellTwo.getRowIndex());
                                       //System.out.println(dummyrow.getCell(l).toString());
                                        }catch(Exception e){
                                        	e.printStackTrace();
                                        continue;
                                    }
                                    }
                        		
                        		
                        	}
                        	
                        	
                        /*	++count;
                        //	int last=sheet2.getLastRowNum();
                        	int last=86;
                        //	System.out.println(last);
                       // 	System.out.println(j);
                        	if(i==4)
                        	{
                        		System.out.println("%%%%%%%%%%%%%%%%");
                        		Row dummyrow1= sheet1.createRow(last+count);
                            	Row Swapingrow1= sheet1.getRow(j);
                            	
                            	for(int l=0;l<=Swapingrow1.getLastCellNum();l++){
                                    try{    Cell cellOne = Swapingrow1.getCell(l);
                                    
                                    
                                        dummyrow1.createCell(l,Cell.CELL_TYPE_BLANK);
                                        Cell cellTwo = dummyrow1.getCell(l);
                                        String dummyStr = cellTwo.toString();
                                        
                                       // System.out.println("*****"+cellOne.toString());
                                        //System.out.println(dummyStr);
                                        
                                        
                                        
                                        //System.out.println("cellOne :: "+cellOne+" cellTwo :: "+cellTwo+" dummyCell :: "+dummyStr);
                                        cellTwo.setCellValue(cellOne.toString());
                                        cellOne.setCellValue(dummyStr);
                                        
                                     //  System.out.println(cellTwo.getRowIndex());
                                       //System.out.println(dummyrow.getCell(l).toString());
                                        }catch(Exception e){
                                        	e.printStackTrace();
                                        continue;
                                    }
                        		
                        	}
                        	}
                        	*/
                        	else{
                        		
                        
                        	Row dummyrow= sheet2.createRow(sheetcount+count);
                        	Row Swapingrow= sheet2.getRow(j);
                        	
                        	for(int l=0;l<=Swapingrow.getLastCellNum();l++){
                                try{    Cell cellOne = Swapingrow.getCell(l);
                                
                                
                                    dummyrow.createCell(l,Cell.CELL_TYPE_BLANK);
                                    dummyrow.getCell(l).setCellValue("N/A");
                                    Cell cellTwo = dummyrow.getCell(l);
                                    String dummyStr = cellTwo.toString();
                                    
                                    
                                    
                                   // System.out.println("*****"+cellOne.toString());
                                    //System.out.println(dummyStr);
                                    
                                    
                                    
                                    //System.out.println("cellOne :: "+cellOne+" cellTwo :: "+cellTwo+" dummyCell :: "+dummyStr);
                                   if(cellOne !=null){
                                    cellTwo.setCellValue(cellOne.toString());
                                    cellOne.setCellValue(dummyStr);
                                   }else
                                   {
                                	   //System.out.println("Coooooo");
                                	 //sheet2.removeRow(Swapingrow);
         
                                	   
                                   }
                                 //  System.out.println(cellTwo.getRowIndex());
                                   //System.out.println(dummyrow.getCell(l).toString());
                                    }catch(Exception e){
                                    	e.printStackTrace();
                                    continue;
                                }
                                }
                        	System.out.println("%%%%%%"+row2.getRowNum());
                        //	sheet2.removeRow(row2);     
                        	
                        	for(int t=0; t<row2.getLastCellNum();t++)
                        	row2.getCell(t).setCellValue("N/A");
                        //	sheet1.shiftRows(j+1, sheet1.getLastRowNum(), -1);
                        //	--j;
                        }
                        }
                    }
                    else
                    {
                    	
                    }
                    
                }catch(Exception e){
                        e.printStackTrace();
                  }
                }
                       
                }
                
               }
        //  }
                
          /*for(int d=0;d<sheet2.getLastRowNum();d++)
          {
        	  if(sheet2.getRow(d) !=null)
        	  {
        		  
        	  }else
        	  {
        		 Row r1= sheet2.createRow(d);
        		  
        		  for(int t=0;t<30;t++)
        		  {
        			  r1.createCell(t).setCellValue("N/A");
        		  }
        		  
        	  }
          }*/
          
          
          
                
        }
        
        count=0;
    writeExcel(workbook, "simple");
      
      
        return workbook;
    }
    
    

    
    
   // public static short cellCounter = 0;
    
    public static XSSFWorkbook createExcel(XSSFWorkbook workbook, int cellCounter, String[] sheetname){
        
        XSSFWorkbook workbookxssfNew = new XSSFWorkbook();
        workbookxssfNew.createSheet("STTM_ONE_Mapping.xlsx");
        XSSFSheet sheetxssfnew = workbookxssfNew.getSheetAt(0);
        XSSFCell cellNew = null;
        XSSFRow rowNew = null;
        
        XLFileReader read=new XLFileReader();
    	String[] header= read.readXLRow("Data_Mapping_Format.xlsx", "Use_Case_1", 1);
    	
        
              
        
        int sheetCount = workbook.getNumberOfSheets();
        try{
              	
        	
        for(int i=sheetCount-1;i>=0 ;i--){
            
            XSSFSheet sheetxssf = workbook.getSheetAt(i);
            int rowCount = sheetxssf.getLastRowNum();
            
            
  
            for(int j=0;j<=rowCount;j++){
            	
            	if(j==0)
            	{
            		
            		if(sheetxssfnew.getRow(j) == null){
                        rowNew = sheetxssfnew.createRow(j);
                    }else{
                        rowNew = sheetxssfnew.getRow(j);
                    }
                    rowNew.setHeight((short) 800);
                    
                    for(int a=0; a<header.length; a++)
                    {
                    	rowNew.createCell(a).setCellValue(header[a]);
                    }
            		
            		
            	 continue;
            	}
                
            	
            	if(i==sheetCount-1)
            	{
            		cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "Target DDM Table");
            		
            		
            	//	System.out.println("@@@@"+cellCounter);
           // 	System.out.println("Sheeeeeetname"+ Sheetname);
            		            		
            		ArrayList<Integer> ll =new ArrayList<>();
            		
            		ll.add(11);
            		ll.add(12);
            		ll.add(13);
            		ll.add(14);
            		ll.add(15);
            		ll.add(16);
            		ll.add(1000);
              		ll.add(17);
            		ll.add(18);
            		ll.add(23);
            		ll.add(19);
            		ll.add(26);
            		ll.add(27);
            		ll.add(21);
            		ll.add(22);
            		ll.add(20);
            		
            		ll.add(1000);
            		ll.add(1000);
            		ll.add(1000);
            		ll.add(1000);
            		
            		if(sheetxssfnew.getRow(j) == null){
                        rowNew = sheetxssfnew.createRow(j);
                       rowNew.createCell(1000).setCellValue("");
                    }else{
                        rowNew = sheetxssfnew.getRow(j);
                       rowNew.createCell(1000).setCellValue("");
                    }
                  //  rowNew.setHeight((short) 800);

		         XSSFRow rowOld;
                        if(sheetxssf.getRow(j) == null){
                        	rowOld = sheetxssf.createRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }else{
                        	rowOld = sheetxssf.getRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }                    

                    short cells = rowOld.getLastCellNum();
                    XSSFCellStyle style = workbookxssfNew.createCellStyle();
                    
                    Font font = workbookxssfNew.createFont();
                	font.setFontHeightInPoints((short)8);
                	font.setFontName("Arial");
                	style.setFont(font);
                    	
                         style.setWrapText(true); 
                         style.setBorderBottom(BorderStyle.THIN);
                         style.setBorderLeft(BorderStyle.THIN);
                         style.setBorderRight(BorderStyle.THIN);
                         style.setBorderTop(BorderStyle.THIN);
                         
                    for(int k=0;k<ll.size();k++){
                    	
      	
                    XSSFCell cellOld = (XSSFCell) rowOld.getCell(ll.get(k)); 
                    
                    //System.out.println(cellOld);
             
                        cellNew = rowNew.createCell(k+cellCounter);
                        cellNew = rowNew.getCell(k+cellCounter);
                        cellNew.setCellStyle(style);
                    if(cellOld != null){
                        copyCell(cellOld, cellNew);
                        //System.out.println(cellNew);
                        //System.out.println(cellCounter);
                    }
                    if(j==0) 
                    sheetxssfnew.autoSizeColumn(k+cellCounter);
                  }
            	}
            	else if (i !=0 && i<sheetCount-1){ 
            		
            		try{
            		System.out.println("DD"+sheetname[i]);

            		
            		if(sheetname[i].toLowerCase().matches("(.*)tpr_to_tmp(.*)"))
            		{
            			System.out.println("tpr_to_tmp"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "TPR to TMP");
            			
            		//	System.out.println("counterrrr"+cellCounter);
            		}
            		else if(sheetname[i].toLowerCase().matches("(.*)tmp_to_ddm(.*)"))
            		{
            		//	System.out.println("TMP_to_DDM"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "TMP to DDM");
            			
            			System.out.println("counterrrr"+cellCounter);
            		}
            		else if(sheetname[i].toLowerCase().matches("(.*)tmp_to_shd(.*)"))
            		{
            		//	System.out.println("TMP_to_SHD"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "TMP to SHD");
            		}
            		else if(sheetname[i].toLowerCase().matches("(.*)shd_to_ddm(.*)") || sheetname[i].toLowerCase().matches("(.*)shd_to_scd(.*)"))
            		{
            			//System.out.println("SHD_to_DDM"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "SHD to DDM");
            		}
            		
            		else if(sheetname[i].toLowerCase().matches("(.*)ddm_to_ddm(.*)"))
            		{
            			//System.out.println("SHD_to_DDM"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "DDM to DDM");
            		}
            		
            		else if(sheetname[i].toLowerCase().matches("(.*)tmp2tmp(.*)"))
            		{
            			//System.out.println("Conf"+sheetname[i]);
            			cellCounter=org.apache.commons.lang.ArrayUtils.indexOf(header, "Conf Table");
            			--cellCounter;
            		//	System.out.println("counterrrr"+cellCounter);
            		}
            		 		
            			
            		
            		
            		ArrayList<Integer> ll =new ArrayList<>();
            		
            		ll.add(1000);
            		ll.add(11);
            		ll.add(12);
            		ll.add(13);
            		ll.add(14);
            		ll.add(15);
            		ll.add(16);
            		ll.add(1000);
               		ll.add(17);
            		ll.add(18);
            		ll.add(23);
            		ll.add(19);
            		ll.add(26);
            		ll.add(27);
            		ll.add(21);
            		ll.add(22);
            		ll.add(20);
            		ll.add(24);
            		ll.add(25);
            		
            		           		
            		            		
            		    if(sheetxssfnew.getRow(j) == null){
                            rowNew = sheetxssfnew.createRow(j);
                        }else{
                            rowNew = sheetxssfnew.getRow(j);
                        }
                     //   rowNew.setHeight((short) 800);
                        
                        XSSFRow rowOld;
                        if(sheetxssf.getRow(j) == null){
                        	rowOld = sheetxssf.createRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }else{
                        	rowOld = sheetxssf.getRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }                    
                        short cells = rowOld.getLastCellNum();
                        XSSFCellStyle style = workbookxssfNew.createCellStyle();
                        
                        Font font = workbookxssfNew.createFont();
                    	font.setFontHeightInPoints((short)8);
                    	font.setFontName("Arial");
                    	style.setFont(font);
                        	
                             style.setWrapText(true); 
                             style.setBorderBottom(BorderStyle.THIN);
                             style.setBorderLeft(BorderStyle.THIN);
                             style.setBorderRight(BorderStyle.THIN);
                             style.setBorderTop(BorderStyle.THIN);
                             
                        for(int k=0;k<ll.size();k++){
                        	
                        XSSFCell cellOld = (XSSFCell) rowOld.getCell(ll.get(k));  
                        //System.out.println(cellOld);
                 
                            cellNew = rowNew.createCell(k+cellCounter);
                            cellNew = rowNew.getCell(k+cellCounter);
                            cellNew.setCellStyle(style);
                        if(cellOld != null){
                            copyCell(cellOld, cellNew);
                            //System.out.println(cellNew);
                            //System.out.println(cellCounter);
                        }
                        if(j==0)
                        sheetxssfnew.autoSizeColumn(k+cellCounter);
                      } 
            		}
            		catch(Exception e){e.printStackTrace();}
                        
            	}
            	
            	else if(i==0){
            		
            		
            		ArrayList<Integer> ll =new ArrayList<>();
            		ll.add(0);
            		ll.add(1000);
            		ll.add(2);
            		
            		for(int a=1000,b=1000; a<1031;a++){
                		ll.add(b);
                		
                	}
            		            		
            		ll.add(5);
            		ll.add(6);
            		ll.add(7);
            		ll.add(8);
            		ll.add(9);
            		ll.add(10);
            		
            		for(int a=1000,b=1000; a<1031;a++){
                		ll.add(b);
                		
                	}
            		
            		ll.add(11);
            		ll.add(12);
            		ll.add(13);
            		ll.add(14);
            		ll.add(15);
            		ll.add(16);
            		ll.add(1000);
               		ll.add(17);
            		ll.add(18);
            		ll.add(23);
            		ll.add(19);
            		ll.add(26);
            		ll.add(27);
            		ll.add(21);
            		ll.add(22);
            		ll.add(20);
            		ll.add(24);
            		ll.add(25);
            		
            		
            		cellCounter=0;
            		
            		//System.out.println("counterss"+cellCounter);
            		
            		            		
            		    if(sheetxssfnew.getRow(j) == null){
                            rowNew = sheetxssfnew.createRow(j);
                        }else{
                            rowNew = sheetxssfnew.getRow(j);
                        }
                     //   rowNew.setHeight((short) 800);
                        
                        XSSFRow rowOld;
                        if(sheetxssf.getRow(j) == null){
                        	rowOld = sheetxssf.createRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }else{
                        	rowOld = sheetxssf.getRow(j);
                        	rowOld.createCell(1000).setCellValue("");
                        }                    
                        short cells = rowOld.getLastCellNum();
                        XSSFCellStyle style = workbookxssfNew.createCellStyle();
                        
                        Font font = workbookxssfNew.createFont();
                    	font.setFontHeightInPoints((short)8);
                    	font.setFontName("Arial");
                    	style.setFont(font);
                        	
                             style.setWrapText(true); 
                             style.setBorderBottom(BorderStyle.THIN);
                             style.setBorderLeft(BorderStyle.THIN);
                             style.setBorderRight(BorderStyle.THIN);
                             style.setBorderTop(BorderStyle.THIN);
                             
                        for(int k=0;k<ll.size();k++){
                        	
                        XSSFCell cellOld = (XSSFCell) rowOld.getCell(ll.get(k));  
                        //System.out.println(cellOld);
                 
                            cellNew = rowNew.createCell(k+cellCounter);
                            cellNew = rowNew.getCell(k+cellCounter);
                            cellNew.setCellStyle(style);
                        if(cellOld != null){
                            copyCell(cellOld, cellNew);
                            //System.out.println(cellNew);
                            //System.out.println(cellCounter);
                        }
                        if(j==0)
                        sheetxssfnew.autoSizeColumn(k+cellCounter);
                      }    
            		
            		
            	}
         }
           // cellCounter = (short) (281+cellCounter);  
            System.out.println("Sheet:"+i);
        }
        }catch(Exception e){
            
        }
       // ++sheetcount;
       // workbookxssfNew=sortCells(workbookxssfNew) ;
        
        
       return workbookxssfNew;
       //writeExcel(workbookxssfNew, sheetname);
    }
    
    
  

	private static void copyCell (XSSFCell cellOld, XSSFCell cellNew){
                //cellNew.setCellStyle(cellOld.getCellStyle());
                //cellNew.setEncoding(cellOld.getEncoding());
         
                switch (cellOld.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                                cellNew.setCellValue(cellOld.getStringCellValue());
                                break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                                cellNew.setCellValue(cellOld.getNumericCellValue());
                                break;
                        case XSSFCell.CELL_TYPE_BLANK:
                                cellNew.setCellValue(HSSFCell.CELL_TYPE_BLANK);
                                break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                                cellNew.setCellValue(cellOld.getBooleanCellValue());
                                break;
                }
        }
    
    
   public static void writeExcel(XSSFWorkbook workbook, String sheetname){
        FileOutputStream out;
        try {
            totalworkbooks[bookcount]="STTM_ONE_Mapping"+sheetname+".xlsx";
          
            ++bookcount;
            out = new FileOutputStream(new File("STTM_ONE_Mapping"+sheetname+".xlsx"));
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException ex) {
         //   Logger.getLogger(WritingExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
         //   Logger.getLogger(WritingExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
			
    }
    
    //for merging different merging excels.
//    public static void MergeFinalExcel()
//    {
//    	XSSFWorkbook workbooksplit=null;
//    	for(int i=0; i<bookcount; i++)
//    	{
//    		File f = new File(totalworkbooks[i]);
//            FileInputStream fis = null;
//            
//            try {
//                fis = new FileInputStream(f);
//                workbooksplit = new XSSFWorkbook(fis);
//                
//                XSSFSheet oldSheet = workbooksplit.getSheetAt(0);
//                XSSFSheet newSheet = workbookfinal.getSheetAt(0);
//                
//                int sheetcount=oldSheet.getLastRowNum();
//                
//                for(int k=0, l=finalrowcount; k<sheetcount && l< finalrowcount+20; k++ ){
//                
//                	if(oldSheet.getRow(k) == null){
//                		Row oldRow = oldSheet.createRow(k);
//                	} else{
//                    Row oldRow = oldSheet.getRow(k);
//                	}
//                	
//                	
//                	// Swap the values here
//                	
//                	
//                
//                
//                }
//                
//                finalrowcount=finalrowcount+sheetcount;
//                
//                
//                
//            } catch (Exception e){e.printStackTrace();}
//            
//            
//            
//    	}
//    }
    
    private static HashMap<String, String> getheaderMap() {
    	
    	HashMap<String, String> hm=new HashMap<>();
    	
  		
  		return hm;
  	}
    
     public static void main(String[] args) {
      //mergeExcel("STTM_Development.xlsx");
     }
}
