package com.excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Merge 
{
	public static Row fixedrow=null;
	public static int pivotcol=0;
   public static void main(String[] args)throws Exception 
   {
	   
	//   for (int i=5;i<50;i++){
		   
	//   File f = new File("E:\\STTM\\DPIM\\Sheet"+i+".xlsx");
	//   if(f.exists() && !f.isDirectory()) { 
		       // do something
		   
	   
	   File file = new File("E:\\STTM\\RMSY\\Sheet1.xlsx");
	   FileInputStream fIP = new FileInputStream(file);
	   XSSFWorkbook workbook = new XSSFWorkbook(fIP);
	   
	   XSSFSheet sheet=workbook.getSheetAt(0);
	   
	   
	   Row row1=null;
	   Row row2=null;
	   Row row11=null;
	   Row row12=null;
	   
	   /*sheet.addMergedRegion(new CellRangeAddress(
			      1, //first row (0-based)
			      14, //last row (0-based)
			      2, //first column (0-based)
			      2 //last column (0-based)
			      ));*/
	   
	  fixedrow= sheet.getRow(1);
	  
	  pivotcol=fixedrow.getLastCellNum();
	  
	  System.out.println("Start");
	  
    //  for(int i=pivotcol; i>=0 ; i--)
    // {
    	  int ccc=0;
    	  int firstrow=0;
    	  int lastrow=0;
    	  
    	  for(int j=1; j<sheet.getLastRowNum();j++)
    	  {
    		  try{
    		  row1=sheet.getRow(j);
    		  row2=sheet.getRow(j+1);
    		  
    		  int sortcol=0;
    		  int prevcol=0;
    		  
    		  if(!row1.getCell(272).toString().equals("N/A"))
    		  {
    			  sortcol=272;
    			  prevcol=271;
    		  }
    		  else if(!row1.getCell(252).toString().equals("N/A"))
    		  {
    			  sortcol=252;
    			  prevcol=251;
    		  }
    		  else if(!row1.getCell(232).toString().equals("N/A"))
    		  {
    			  sortcol=232;
    			  prevcol=231;
    		  }
    		    else if(!row1.getCell(212).toString().equals("N/A"))
    		  {
    			  sortcol=212;
    			  prevcol=211;
    		  }
    		  
    		  else if(!row1.getCell(192).toString().equals("N/A"))
    		  {
    			  sortcol=192;
    			  prevcol=191;
    		  }
    		  else if(!row1.getCell(172).toString().equals("N/A"))
    		  {
    			  sortcol=172;
    			  prevcol=171;
    		  }
    		  /*
    		  if(!row1.getCell(152).toString().equals("N/A"))
    		  {
    			  sortcol=152;
    			  prevcol=151;
    		  }
    		  else if(!row1.getCell(132).toString().equals("N/A"))
    		  {
    			  sortcol=132;
    			  prevcol=131;
    		  }
    		  else if(!row1.getCell(112).toString().equals("N/A"))
    		  {
    			  sortcol=112;
    			  prevcol=111;
    		  }
    		  else if(!row1.getCell(92).toString().equals("N/A"))
    		  {
    			  sortcol=92;
    			  prevcol=91;
    		  }
    		  else if(!row1.getCell(72).toString().equals("N/A"))
    		  {
    			  sortcol=72;
    			  prevcol=71;
    		  }
    		  else if(!row1.getCell(53).toString().equals("N/A"))
    		  {
    			  sortcol=53;
    			  prevcol=52;
    		  }*/
    		  
    		  
    		  Cell c1=row1.getCell(sortcol);
    		  Cell c2=row2.getCell(sortcol);
    		  
    		  Cell pc1=row1.getCell(prevcol);
    		  Cell pc2=row2.getCell(prevcol);
    		  
    		  String cell1;
    		  String cell2;
    		  
    		 if(pc1.toString().contains(","))
    		  {
    			  cell1= pc1.toString().split(",")[0];
    		  }
    		  else
    		  {
    			  cell1=pc1.toString();
    		  }
    		  
    		  if(pc2.toString().contains(","))
    		  {
    			  cell2= pc2.toString().split(",")[0];
    		  }
    		  else
    		  {
    			  cell2=pc2.toString();
    		  }
    		  
    		  
    	//	  if(row1 !=null ||row2 !=null){
    		  
    		  if(cell1.trim().equals(cell2.trim())&&(c1.toString().trim().equals(c2.toString().trim())) && (!(cell1.equals("N/A")) && !(cell2.equals("N/A")))){
    			////system.out.println("Here");
       		  if(ccc==0)
    		  {
       			  ccc=1;
    			 firstrow=row1.getRowNum(); 
    		  }
       		     lastrow=row2.getRowNum();
    		  }
    		  else if(!(cell1.equals("N/A")) && !(cell2.equals("N/A")))
    		  {
    			  ccc=0;
    			  
    			  // Below loops need to be done in a single for loop in future.
    			  
    			  for(int k=270; k<pivotcol;k++){
    				  
    				   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    			
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(272);
    	    		  Cell picell2=row12.getCell(272);
    	    		  
    	    		  Cell precell1=row11.getCell(271);
    	    		  Cell precell2=row12.getCell(271);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  
    	    		  
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else 
    	    		  {
    	    			  cc1=0;
    	    			//  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  }
    			  
    			  for(int k=250; k<270;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(252);
    	    		  Cell picell2=row12.getCell(252);
    	    		  
    	    		  Cell precell1=row11.getCell(251);
    	    		  Cell precell2=row12.getCell(251);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  
    	    		  
    	    		  if(c11.toString().contains(",") && c11.toString().split(",").length >0)
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(",") && c12.toString().split(",").length >0)
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  if(k==264){
    	    		  System.out.println("cell1:"+incell1);
    	    		  System.out.println("cell2:"+incell2);}
    	    		
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  for(int k=230; k<250;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(232);
    	    		  Cell picell2=row12.getCell(232);
    	    		  
    	    		  Cell precell1=row11.getCell(231);
    	    		  Cell precell2=row12.getCell(231);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  /* */
    	    		  Cell pMergepicell1=row11.getCell(252);
    	    		  Cell pMergepicell2=row12.getCell(252);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(251);
    	    		  Cell pMergeprecell2=row12.getCell(251);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  for(int k=210; k<230;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(212);
    	    		  Cell picell2=row12.getCell(212);
    	    		  
    	    		  Cell precell1=row11.getCell(211);
    	    		  Cell precell2=row12.getCell(211);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(232);
    	    		  Cell pMergepicell2=row12.getCell(232);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(231);
    	    		  Cell pMergeprecell2=row12.getCell(231);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	   if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  for(int k=190; k<210;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(192);
    	    		  Cell picell2=row12.getCell(192);
    	    		  
    	    		  Cell precell1=row11.getCell(191);
    	    		  Cell precell2=row12.getCell(191);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(212);
    	    		  Cell pMergepicell2=row12.getCell(212);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(211);
    	    		  Cell pMergeprecell2=row12.getCell(211);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  for(int k=170; k<190;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(172);
    	    		  Cell picell2=row12.getCell(172);
    	    		  
    	    		  Cell precell1=row11.getCell(171);
    	    		  Cell precell2=row12.getCell(171);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(192);
    	    		  Cell pMergepicell2=row12.getCell(192);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(191);
    	    		  Cell pMergeprecell2=row12.getCell(191);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  for(int k=150; k<170;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(152);
    	    		  Cell picell2=row12.getCell(152);
    	    		  
    	    		  Cell precell1=row11.getCell(151);
    	    		  Cell precell2=row12.getCell(151);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(172);
    	    		  Cell pMergepicell2=row12.getCell(172);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(171);
    	    		  Cell pMergeprecell2=row12.getCell(171);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&((picell1.toString().equals(picell2.toString()))) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  
    			  for(int k=130; k<150;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(132);
    	    		  Cell picell2=row12.getCell(132);
    	    		  
    	    		  Cell precell1=row11.getCell(131);
    	    		  Cell precell2=row12.getCell(131);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(152);
    	    		  Cell pMergepicell2=row12.getCell(152);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(151);
    	    		  Cell pMergeprecell2=row12.getCell(151);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
   
    			  
    			  for(int k=110; k<130;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(112);
    	    		  Cell picell2=row12.getCell(112);
    	    		  
    	    		  Cell precell1=row11.getCell(111);
    	    		  Cell precell2=row12.getCell(111);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(132);
    	    		  Cell pMergepicell2=row12.getCell(132);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(131);
    	    		  Cell pMergeprecell2=row12.getCell(131);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  
    			  for(int k=90; k<110;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(92);
    	    		  Cell picell2=row12.getCell(92);
    	    		  
    	    		  Cell precell1=row11.getCell(91);
    	    		  Cell precell2=row12.getCell(91);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(112);
    	    		  Cell pMergepicell2=row12.getCell(112);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(111);
    	    		  Cell pMergeprecell2=row12.getCell(111);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }

    			  
    			  for(int k=70; k<90;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(72);
    	    		  Cell picell2=row12.getCell(72);
    	    		 
    	    		  Cell precell1=row11.getCell(71);
    	    		  Cell precell2=row12.getCell(71);
    	    		 
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(92);
    	    		  Cell pMergepicell2=row12.getCell(92);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(91);
    	    		  Cell pMergeprecell2=row12.getCell(91);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  for(int k=51; k<70;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(52);
    	    		  Cell picell2=row12.getCell(52);
    	    		  
    	    		  Cell precell1=row11.getCell(51);
    	    		  Cell precell2=row12.getCell(51);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(72);
    	    		  Cell pMergepicell2=row12.getCell(72);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(71);
    	    		  Cell pMergeprecell2=row12.getCell(71);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
    			  
    			  
    			  for(int k=0; k<51;k++)
    			  {  				  
	   				  
    				  int firstrow1=0;
    				  int lastrow1=0;
    				  int cc1=0;
    				//  //system.out.println("fir"+firstrow);
    			//	  //system.out.println("Last"+lastrow);
    			  for(int t= firstrow ;t<=lastrow;t++)
    			  {
    				  try{
      				  row11=sheet.getRow(t);
    				  row12=sheet.getRow(t+1);
    	    		  
    	    		  Cell c11=row11.getCell(k);
    	    		  Cell c12=row12.getCell(k);
    	    		  
    	    		  Cell picell1=row11.getCell(35);
    	    		  Cell picell2=row12.getCell(35);
    	    		  
    	    		  Cell precell1=row11.getCell(34);
    	    		  Cell precell2=row12.getCell(34);
    	    		  
    	    		  String pincell1=null;
    	    		  String pincell2=null;
    	    		  
    	    		  if(precell1.toString().contains(","))
    	    		  {
    	    			  pincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell1=precell1.toString();
    	    		  }
    	    		  
    	    		  if(precell2.toString().contains(","))
    	    		  {
    	    			  pincell2= precell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pincell2=precell2.toString();
    	    		  }
    	    		  
    	    		  String incell1;
    	    		  String incell2;
    	    		  
    	    		  if(c11.toString().contains(","))
    	    		  {
    	    			  incell1= c11.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell1=c11.toString();
    	    		  }
    	    		  
    	    		  if(c12.toString().contains(","))
    	    		  {
    	    			  incell2= c12.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  incell2=c12.toString();
    	    		  }
    	    		  
    	    		  
    	    		  Cell pMergepicell1=row11.getCell(52);
    	    		  Cell pMergepicell2=row12.getCell(52);
    	    		  
    	    		  Cell pMergeprecell1=row11.getCell(51);
    	    		  Cell pMergeprecell2=row12.getCell(51);
    	    		  
    	    		  String pMergepincell1=null;
    	    		  String pMergepincell2=null;
    	    		  
    	    		  if(pMergeprecell1.toString().contains(","))
    	    		  {
    	    			  pMergepincell1= precell1.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell1=pMergeprecell1.toString();
    	    		  }
    	    		  
    	    		  if(pMergeprecell2.toString().contains(","))
    	    		  {
    	    			  pMergepincell2= pMergeprecell2.toString().split(",")[0];
    	    		  }
    	    		  else
    	    		  {
    	    			  pMergepincell2=pMergeprecell2.toString();
    	    		  } 
    	    		 
    	    		   
    	    		  
    	    		  
    	    		  
    	    		
    	    		  if(incell1.trim().equals(incell2.trim()) && row12.getRowNum() !=lastrow+1 &&(picell1.toString().equals(picell2.toString())) && ((pMergepicell1.toString().equals(pMergepicell2.toString()))&& (pMergepincell1.toString().equals(pMergepincell2.toString()))&&(!pMergepicell1.equals("N/A")||!pMergepicell2.equals("N/A"))) && (pincell1.toString().equals(pincell2.toString())) ){
    	         	  if(cc1==0)
    	      		  {
    	         			  cc1=1;
    	         			  firstrow1=row11.getRowNum(); 
    	      		  }
    	         		     lastrow1=row12.getRowNum();
    	      		  }
    	    		  else
    	    		  {
    	    			  cc1=0;
    	    			  //system.out.println("Merging:"+firstrow1+"to"+lastrow1+"**"+c11.toString());
    	    			  sheet.addMergedRegion(new CellRangeAddress(
    	    					  firstrow1, //first row (0-based)
    	    					  lastrow1, //last row (0-based)
    	    				      k, //first column (0-based)
    	    				      k //last column (0-based)
    	    					  ));
    	    		  }
    				  
    	    		  }catch(Exception e){
    	    			  //e.printStackTrace();
    	    			  continue;
    	    		  }
    				  
    			  }
    			  
    			  
    			  }
   			  
    		  }
    		  }catch(Exception e){ 
    		  e.printStackTrace();
    		  }
    	//	  }
    	  }
    	  
 //   }
	   
	   // writing to o/p stream
	   //system.out.println("Writing");
	   FileOutputStream out = new FileOutputStream( 
			      new File("E:\\STTM\\RMSY\\1_test.xlsx"));
			      workbook.write(out);
			     
			      out.close();
			      
			      
	//   }
	//   System.gc();
			      
	//   }
	   
   }
}