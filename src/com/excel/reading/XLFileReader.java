package com.excel.reading;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.apache.poi.ss.usermodel.DateUtil;


//import seetestClient.MyClient;

//import exception.ElementNotFoundException;

public class XLFileReader {
	
	
	static Logger log = Logger.getLogger(XLFileReader.class.getName());
	FormulaEvaluator formulaEval;
	
	//MyClient client;
	
//	public XLFileReader(MyClient myclient)
//	{
//		client = myclient;
//	}
	public XLFileReader()
	{
		
	}
	
	/* 
	 * Method Name: getWorkSheet()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 * 
	 * Output Parameters : sheet                (Sheet)
	 * 
	 * Description		: This method takes XL file name and Sheet name as input and returns the Sheet object.
	 * 
*/
		
public Sheet getWorkSheet(String FileName, String SheetName)
{
	
	Sheet SheetObject = null;
    Workbook workbook = null;
	try
	{
		FileInputStream fis = new FileInputStream(FileName);
		
        //Create Workbook instance for xlsx/xls/xlsm file input stream
    	        
        if((FileName.toLowerCase().endsWith("xlsx"))||((FileName.toLowerCase().endsWith("xlsm"))))
        {
            workbook = new XSSFWorkbook(fis);
            formulaEval = workbook.getCreationHelper().createFormulaEvaluator();  
        }
        else if(FileName.toLowerCase().endsWith("xls"))
        {
            workbook = new HSSFWorkbook(fis);
        }
		
    	
        SheetObject = workbook.getSheet(SheetName);
        
        if(SheetObject == null)
        {
        	throw new Exception("The Sheet "+SheetName+ " is not found in the File "+FileName+"\n\n");
        }
	}
	catch(Exception e)
	{
		log.error("Exception Occurred",e);
		e.printStackTrace();
	}
	
	return SheetObject;
}

/**************************************************************************************************************************************************************************/
/**************************************************************************************************************************************************************************/

	
/* 
 * Method Name: readHeaderIndex()
 * 
 * Input Parameters  : File Name 			(String)
 *  				 : Sheet Name           (String)
 *  				 : Header               (String)
 * 
 * Output Parameters : HeaderIndex          (integer)
 * 
 * Description		: This method takes XL file name, Sheet name and header as input and returns index of the Header file.
 * 
*/	
	
	public int readHeaderIndex(String fileName,String sheetName, String Header)
    {
		String Parameters = "FileName : "+fileName+", SheetName : "+sheetName+", Header : "+Header+"\n";
		
		log.info("Inside Method readHeaderIndex() \n Parameters : "+Parameters);

		int ColumnIndex	=	0;
		boolean found = false;

    	try
    	{
	    	
	    	Sheet sheet = getWorkSheet(fileName,sheetName);

	    	if(sheet != null)
	    	{		
		    	Row HeaderRow=sheet.getRow(0);
		    	
		    	Iterator<Cell> HeaderCellIterator = HeaderRow.cellIterator();
		    	
		    	while(HeaderCellIterator.hasNext())
		    	{
		    		Cell celltoFind = HeaderCellIterator.next();
		    		if(celltoFind.toString().equals(Header))
		    		{
		    			ColumnIndex		= 	celltoFind.getColumnIndex();
		    			found = true;
		    		}
		    	}
		    	
		    	if(found == false)
		    	{
		    		throw new Exception("The Column "+Header+" is not found in the "+sheetName+" sheet of the File "+fileName);
		    	}
	    	}
    	}catch(Exception e)
    	{
    		e.printStackTrace();
    		log.error("Exception Occurred!!!!", e);
    	}
    	
    	
    	return ColumnIndex+1;
    	
    }
	
	/**************************************************************************************************************************************************************************/
	/**************************************************************************************************************************************************************************/

		
	/* 
	 * Method Name: readXLatIndex()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 *  				 : Row to Read          (integer)
	 *  				 : Column to Read       (integer)
	 * 
	 * Output Parameters : Cell Value           (String)
	 * 
	 * Description		: This method takes XL file name, Sheet name, Row index and Column index and returns the Cell Value.
	 * 
	*/	

	


	
	
		
	public String readXLatIndex(String fileName,String sheetName, int RowIndex, int ColumnIndex)
	{
		String Parameters = "FileName : "+fileName+", SheetName : "+sheetName+", Row Index : "+RowIndex+", Column Index : "+ColumnIndex+"\n";
		log.info("Inside Method readXLatIndex() \n Parameters : "+Parameters);
		
		RowIndex-= 1;
		ColumnIndex-=1;
		
		String ColumnStringValue	= "";
		
		
		try
		{
	    	Sheet sheet = getWorkSheet(fileName,sheetName);
	    	
    	   if(sheet!=null)
    	   {
		     	Row DataRow=sheet.getRow(RowIndex);
		    	
		    	Cell celltoFind = DataRow.getCell(ColumnIndex);
		    	
		    	if(celltoFind != null)
		    	{
		    	
			    	switch( celltoFind.getCellType())
			    	{
			    		case Cell.CELL_TYPE_BLANK:
					    	ColumnStringValue = "";
					    	
			    		case Cell.CELL_TYPE_NUMERIC:
			    			
			    			//ColumnStringValue = celltoFind.toString().replaceAll("\\.?0*$", "");   //Removing Trailing Zeros and assigning to the String
			    			//ColumnStringValue = celltoFind.toString().replaceAll("\\.?0*$", "");
			    			long i = (long)celltoFind.getNumericCellValue();//Getting Numeric value from the sheet and Type casting to 'long' type to hold more than 12 digits 
			    			ColumnStringValue = String.valueOf(i);  
			    			break;
				    	
			    		case Cell.CELL_TYPE_STRING:
			    			
				    		ColumnStringValue = celltoFind.toString();
				    		break;
			    		
			    		case Cell.CELL_TYPE_BOOLEAN:
			    			
				    		ColumnStringValue = celltoFind.toString();
				    		
				    		break;
				    		
			    		case Cell.CELL_TYPE_ERROR:
			    			
				    		ColumnStringValue = celltoFind.toString();
				    		break;
			    		
			    		case Cell.CELL_TYPE_FORMULA:
			    			
				    		ColumnStringValue = celltoFind.toString();
				    		
				    		
				    		break;
			    		
				    	default:
				    		System.out.println("Default");
					    		ColumnStringValue	=	"";
						    	
				    	}
			    	}
    	   		}
	    	}
			catch(Exception e)
	    	{
	    		e.printStackTrace();
	    	}
	    	return ColumnStringValue;
	    	
	    }

	/**************************************************************************************************************************************************************************/
	/**************************************************************************************************************************************************************************/


	/* 
	 * Method Name: readXLRow()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 *  				 : Row to read          (integer)
	 * 
	 * Output Parameters : XLRow                (String[])
	 * 
	 * Description		: This method returns the XL Row in an array of string.
	 * 
	*/

	public String[] readXLRow(String fileName,String SheetName,int RowtoRead)
	{
		String Parameters = "FileName : "+fileName+", SheetName : "+SheetName+", Row to Read : "+RowtoRead+"\n";
		log.info("Inside Method readXLRow() \n Parameters : "+Parameters);	
		
		RowtoRead = RowtoRead-1;
		String[] cellsinRow = null;
		
		try
		{
			Sheet sheet = getWorkSheet(fileName,SheetName);
	    	
			if(sheet!=null)
			{
		    	int i=0;
	    	
		    	int cellCount = readNumberofCellsinXL(fileName, SheetName);
                 System.out.println("CellCount"+cellCount);
		    	
		    	cellsinRow	=	new String[cellCount];

		    	Row DataRow=sheet.getRow(RowtoRead);
		    	

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
			}
		
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
		return cellsinRow;
	}


	/* 
	 * Method Name: readNumberofRowsinXL()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 * 
	 * Output Parameters : number of Rows       (integer)
	 * 
	 * Description		: This method returns the number of rows in XL.
	 * 
	*/


	public int readNumberofRowsinXL(String fileName,String SheetName)
	{
		String Parameters = "FileName : "+fileName+", SheetName : "+SheetName+"\n";
		log.info("Inside Method readNumberofRowsinXL() \n Parameters : "+Parameters);	
		
		int NumberofRows = 0;
		
		
		try
		{
	   		Sheet sheet = getWorkSheet(fileName,SheetName);
	   		if(sheet!=null)
	   		{
	   			NumberofRows = sheet.getLastRowNum()+1;
	   		}
	    	
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		    	
		return NumberofRows;
	}
	
	/**************************************************************************************************************************************************************************/
	/**************************************************************************************************************************************************************************/

	/* 
	 * Method Name: readNumberofCellsinXL()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 * 
	 * Output Parameters : number of Cells      (integer)
	 * 
	 * Description		: This method returns the number of cells in XL.
	 * 
	*/

	public int readNumberofCellsinXL(String fileName,String SheetName)
	{
		String Parameters = "FileName : "+fileName+", SheetName : "+SheetName+"\n";	
		log.info("Inside Method readNumberofCellsinXL() \n Parameters : "+Parameters);	
		
		int cellCount = 0;
		
		try
		{
	   		Sheet sheet = getWorkSheet(fileName,SheetName);
	   		if(sheet!=null)
	   		{
		   		Row DummyRow=sheet.getRow(0);
	
				Iterator<Cell> HeaderCellIterator = DummyRow.cellIterator();
		    	
				// Iterating to find the cell count
		    	
		    	while(HeaderCellIterator.hasNext())
		    	{
		    		Cell CurrentRowCells = HeaderCellIterator.next();
		    		cellCount++;
		    	}
	   		}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		

		return cellCount;
	}
	
	/**************************************************************************************************************************************************************************/
	/**************************************************************************************************************************************************************************/
		
		
	/* 
	 * Method Name: writeToXLCell()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : Sheet Name           (String)
	 *  				 : row 					(Integer)
	 *  				 : Column				(Integer)
	 *  				 : Value to Write       (String)
	 * 
	 * 
	 * Description		: This method writes a String to the Given Cell.
	 * 
	*/

	public void writeToXLCell(String FileName, String SheetName, int row, int column, String ValuetoWrite, String result)throws Exception
	{
		String Parameters = "FileName : "+FileName+", SheetName : "+SheetName+", Row Number : "+row+", Column Number : "+column+", \n Value to Write : " +ValuetoWrite+"\n";
		log.info("Inside Method writeToXLCell() \n Parameters : "+Parameters);	
		try
		{
			Sheet sheet = getWorkSheet(FileName,SheetName);
			Workbook workbook = sheet.getWorkbook();
            /* Get access to XSSFCellStyle */
            XSSFCellStyle my_style = (XSSFCellStyle) workbook.createCellStyle();
            
            /* First, let us draw a thick border so that the color is visible */            
            my_style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);             
            my_style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);            
            my_style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);              
            my_style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
            
            my_style.setBorderColor(BorderSide.LEFT,new XSSFColor(Color.black));
            my_style.setBorderColor(BorderSide.RIGHT,new XSSFColor(Color.black));
            my_style.setBorderColor(BorderSide.TOP,new XSSFColor(Color.black));
            my_style.setBorderColor(BorderSide.BOTTOM,new XSSFColor(Color.black));
            
            if(result.equals("pass")){
            my_style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            my_style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
            }else if (result.equals("fail")){
            my_style.setFillForegroundColor(IndexedColors.RED.getIndex());
            my_style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
            }
            else
            {
            	my_style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
                my_style.setFillPattern(CellStyle.SOLID_FOREGROUND);
            }
            
			
			row-=1;
			column-=1;		
			
	  		
	  		if(sheet!=null)
	  		{
		  		Row rowObject=sheet.getRow(row);
		    	Cell cell=rowObject.createCell(column);
		        cell.setCellValue(ValuetoWrite);
		        cell.setCellStyle(my_style);
		        
		        FileOutputStream outFile =new FileOutputStream(new File(FileName));
		        
		       
		        workbook.write(outFile);
		        outFile.close();
	  		}
		}
		catch(Exception e)
		{		
			e.printStackTrace();
		}
		
	}
	    	
	/**************************************************************************************************************************************************************************/
	/**************************************************************************************************************************************************************************/
		
		
	/* 
	 * Method Name: saveCopy()
	 * 
	 * Input Parameters  : File Name 			(String)
	 *  				 : New File Name           (String)
	 * 
	 * 
	 * Description		: This method Saves a copy of the given file with a new name given
	*/
	   
	    public void saveCopy(String FileName, String NewFileName)
	    {
	    	String Parameters = "Source FileName : "+FileName+", Destination FileName : "+NewFileName+"\n";	
	    	log.info("Inside Method saveCopy() \n Parameters : "+Parameters);	
	    	
	    	try
	    	{
	    	   	 File source	=	new File(FileName);
	    	     File dest		=	new File(NewFileName); 
	    	     
	    	     Files.copy(source.toPath(), dest.toPath());
	    	     
	     	}
	    	catch(Exception e)
	    	{
	    		e.printStackTrace();
	    	}
	    }
	    
	    
	}



