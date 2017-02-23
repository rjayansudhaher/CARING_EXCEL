/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.excel.reading;

import static com.excel.reading.ExcelMergerNew.bookcount;
import static com.excel.reading.ExcelMergerNew.totalworkbooks;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author PR355016
 */
public class ExcelFolderReader {
    
    public static XSSFWorkbook workbookProp = null;
    public static String mappingRowNo = "";
    
    public static void readMappingName(String name){
        
        File f = new File(name);
        FileInputStream fis = null;
        
        try {
            fis = new FileInputStream(f);
            workbookProp = new XSSFWorkbook(fis);
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelFolderReader.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelFolderReader.class.getName()).log(Level.SEVERE, null, ex);
        }finally{
            try {
                if(fis!=null){
                fis.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(ExcelFolderReader.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        }
        
        XSSFSheet propSheet = workbookProp.getSheetAt(0);
        for(int i=0;i<=propSheet.getLastRowNum();i++){
            XSSFRow row = propSheet.getRow(i);
            XSSFCell cell = null;
            
           boolean exists=false;
            
           for (int j=0; j<row.getLastCellNum(); j++)
           {
        	 if(row.getCell(j) == null)
        	 {
        		 continue;
        	 }
            cell = row.getCell(j);
           
            if(readFolder(cell.toString()))
            {
            	exists=true;
            	break;
            }
            
           }
            try{
            	
            if(exists){
                if(mappingRowNo.trim().equalsIgnoreCase("")){
                    mappingRowNo = String.valueOf(i);

                }else{
                    mappingRowNo = mappingRowNo+","+i;
                }
            }
            }catch(Exception e){
                
            }
        }    
    }
    
    public static boolean readFolder(String name){
    	
    	boolean ret = false;
    	if(name.contains("/"))
        {
        	String[] a= name.split("/");
        	for (int i=0; i<a.length;i++)
        	{
        		File f = new File("./input/PEPA/"+a[i].trim()+".xlsx");
                //System.out.println(f.getAbsolutePath());
                //System.out.println(f.exists());
                if(f.exists()){
                	ret= true;
                }else{
                	ret= false;
                }
        	}
        	
        }
    	else{
        File f = new File("./input/PEPA/"+name.trim()+".xlsx");
        if(f.exists()){
        	ret= true;
        }else{
        	ret= false;
        }
    	}
    	return ret;
    }
    
    
    public static void readMappingFile(){
        ArrayList<XSSFWorkbook> workbookList = new ArrayList<XSSFWorkbook>();
        String rowNum[] = mappingRowNo.split(",");
        
        System.out.println("row num"+rowNum.length);
        
        
        
        TEST:
        for(int i=0;i<rowNum.length;i++){
        	
        	int cellCounter=0;
        	int slash=0;
        	int rowie=1;
        	XSSFWorkbook dummyWorkbook = new XSSFWorkbook();
        	System.out.println("##"+rowNum[i]);
        	int sheetcount=0;
        	
        	Integer.parseInt(rowNum[i]);
            XSSFRow row = workbookProp.getSheetAt(0).getRow(Integer.parseInt(rowNum[i]));
            String[] Sheetname=new String[row.getLastCellNum()];
           
            for(int j=row.getLastCellNum();j>=0;j--){
            	
            	if(row.getCell(j)==null){
            		continue;
            	}    
            	
            	String[] split={row.getCell(j).toString()}; 	
            	String[] srcsplit={row.getCell(j).toString()};
            	slash=0;
           if(row.getCell(j).toString().contains("/"))
           {
        	   split= row.getCell(j).toString().split("/");
        	//   srcsplit= row.getCell(j).toString().split("/");
        	 //  ++slash;
        	   sheetcount++;
        	   
           }
           rowie=1;
               	   
             for(int k=0;k<split.length;k++){
            	 System.out.println(split[k].toString());
            	
            	 if(row.getCell(j).toString().contains("/"))
                 {
              	 //  split= row.getCell(j).toString().split("/");
              	   srcsplit= row.getCell(j).toString().split("/");
              	   ++slash;
              	   
                 }
            	
             if(readFolder(split[k].toString())){
                 File f = new File("./input/PEPA/"+split[k].toString().trim()+".xlsx");
                 
                 if(split[k].toString() !=null){
                 Sheetname[j]= split[k].toString();
                 }
                 System.out.println("sssssssssssssssss"+Sheetname[j]);
                 try {
                         	 
                     XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(f));
                     XSSFSheet sheet = work.getSheetAt(0);
                     //ExcelMergerNew.cellCompare(work);
                     //int sheetNum = dummyWorkbook.getNumberOfSheets();
                     //System.out.println(dummyWorkbook.getNumberOfSheets());
                     if(slash==0){
                     XSSFSheet newsheet = dummyWorkbook.createSheet();
                     for(int l=0;l<=sheet.getLastRowNum();l++){
                         XSSFRow oldRow = sheet.getRow(l);
                         XSSFRow newRow = newsheet.createRow(l);
                         for(int m=0;m<=oldRow.getLastCellNum();m++){
                             try{
                             XSSFCell oldCell = oldRow.getCell(m);
                             XSSFCell newCell = newRow.createCell(m);
                             newCell.setCellValue(oldCell.toString());
                             }catch(Exception e){
                                 XSSFCell newCell = newRow.createCell(m);
                             newCell.setCellValue("");
                             }
                         }
                     }
                    }else
                    {
                    	 XSSFSheet newsheet;
                    	if(slash==1)
                    	{
                    		newsheet= dummyWorkbook.createSheet("mergesheet"+sheetcount+"");
                    	 }
                    	else
                    	{
                    	 newsheet = dummyWorkbook.getSheet("mergesheet"+sheetcount+"");
                    	}
                                             	
                    	for(int l=1;l<=sheet.getLastRowNum();l++){
                            XSSFRow oldRow = sheet.getRow(l);
                            XSSFRow newRow = newsheet.createRow(rowie);
                            for(int m=0;m<=oldRow.getLastCellNum();m++){
                                try{
                                XSSFCell oldCell = oldRow.getCell(m);
                                XSSFCell newCell = newRow.createCell(m);
                                newCell.setCellValue(oldCell.toString());
                                
                                }catch(Exception e){
                                    XSSFCell newCell = newRow.createCell(m);
                                newCell.setCellValue("");
                                }
                            }
                            ++rowie;
                        }
                    }
                     
                 } catch (FileNotFoundException ex) {
                     Logger.getLogger(ExcelFolderReader.class.getName()).log(Level.SEVERE, null, ex);
                 } catch (IOException ex) {
                     Logger.getLogger(ExcelFolderReader.class.getName()).log(Level.SEVERE, null, ex);
                 }
                
             }else{
                 System.out.println("one of the mapping file not found."+split[k].toString());
                 continue ; //TEST;
             }  
            }
             
            }
            //System.out.println("sheetName::" +Sheetname);
          //writeExcel(dummyWorkbook, "ASD");
            Sheetname=reverse(Sheetname);
            XSSFWorkbook workbookNew = ExcelMergerNew.mergeExcel(dummyWorkbook, cellCounter, Sheetname);
            workbookList.add(workbookNew);
            //System.out.println("list of workbooks are : "+workbookList.size());
            //writeExcel(workbookList.get(0), "firstSheet");
            //writeExcel(workbookList.get(1), "secondSheet");
            mergeFinalExcel(workbookList);
           
        }
     //  System.out.println(dummyWorkbook.getNumberOfSheets());   
       // writeExcel(dummyWorkbook);
       
    // ExcelMergerNew.mergeExcel(dummyWorkbook);
        
        
        
    }
    
    public static String[] reverse(String[] a){

    	String[] reversedArray = new String[a.length];

    	for(int i = 0 ; i<a.length; i++){
    	reversedArray[i] = a[a.length -1 -i];


    	}
		return reversedArray;
    }
    
    public static void writeExcel(XSSFWorkbook workbook, String sheetname){
        FileOutputStream out;
        try {
            out = new FileOutputStream(new File(sheetname+".xlsx"));
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException ex) {
         //   Logger.getLogger(WritingExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
         //   Logger.getLogger(WritingExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
			
    }
    
    public static void mergeFinalExcel(ArrayList<XSSFWorkbook> workbookList){
        XSSFWorkbook finalWorkbook = new XSSFWorkbook();
        XSSFSheet newSheet = finalWorkbook.createSheet("JHRY");
        
        int mergeCounter = 0;
        
        for(int i=0;i<workbookList.size();i++){
            XSSFWorkbook dum = workbookList.get(i);
            XSSFSheet oldSheet = dum.getSheetAt(0);
            
            //System.out.println("workbook row size is : "+oldSheet.getLastRowNum());
            XSSFCellStyle style = null;
            if(i == 0){
                XSSFRow row = newSheet.createRow(0);
                for(int l=0;l<oldSheet.getRow(0).getLastCellNum();l++){
                    
                    if(l==0 || l==3 || l==6 || l==20|| l==34 ||l==51 || l==70|| l==90||l==110||l==130||l==150||l==170||l==190||l==210||l==230||l==249||l==269||l==281||l==298||l==308){
                        style = finalWorkbook.createCellStyle();
                    
                        Font font = finalWorkbook.createFont();
                        font.setFontHeightInPoints((short)8);
                        font.setFontName("Arial");
                        style.setFont(font);
                        Random rand = new Random(); 
                        int L = 50;
                        int H = 255;
                        int r = rand.nextInt(H-L)+L;
                        int g = rand.nextInt(H-L)+L;
                        int b = rand.nextInt(H-L)+L;
                        byte[] rgb = {(byte)r,(byte)g,(byte)b};
                        style.setFillForegroundColor(new XSSFColor(rgb));
                        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                        style.setWrapText(true); 
                        style.setBorderBottom(BorderStyle.THIN);
                        style.setBorderLeft(BorderStyle.THIN);
                        style.setBorderRight(BorderStyle.THIN);
                        style.setBorderTop(BorderStyle.THIN);
                        }
                    
                XSSFCell cell = row.createCell(l);    
                copyCell(oldSheet.getRow(0).getCell(l), cell);
                if(style != null){
                        cell.setCellStyle(style);
                    }
                }
                
                
            }
            for(int j=1;j<=oldSheet.getLastRowNum();j++){
                XSSFRow row = oldSheet.getRow(j);
                XSSFRow newRow = newSheet.createRow(j+mergeCounter);
                for(int k=0;k<row.getLastCellNum();k++){
                    XSSFCell oldCell = row.getCell(k);
                    
                    XSSFCell newCell = newRow.createCell(k);
                    try{
                    	
                    newCell.setCellValue(oldCell.getStringCellValue());
                    }catch(Exception e){

                    }
                    
                }
            }
            mergeCounter = mergeCounter + oldSheet.getLastRowNum();
            //System.out.println("merger counter is :"+mergeCounter);
        }
        
        //System.out.println("new sheet length is : " + newSheet.getLastRowNum());
        
        
        writeExcel(finalWorkbook, "UNFI");
        
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
    
    public static void main(String args[]){
        readMappingName("Data Lineage RICA.XLSX");
        System.out.println(mappingRowNo);
        readMappingFile();
    }
}
