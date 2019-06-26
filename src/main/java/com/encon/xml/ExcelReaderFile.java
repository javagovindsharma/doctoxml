package com.encon.xml;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;


public class ExcelReaderFile {
    
         String FILE_NAME="HTML code for letters with accents V2.xlsx";
  
          public   Map<String,String> readExcelFile(){
              Map<String,String> maps=new HashMap<String,String>();
              try {
        	  String path = System.getProperty("user.dir");
        	  
                  FileInputStream excelFile = new FileInputStream(new File(path+"/"+FILE_NAME));
                  @SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(excelFile);
                  Sheet datatypeSheet = workbook.getSheetAt(0);
                  Iterator<Row> iterator = datatypeSheet.iterator();

                  while (iterator.hasNext()) {

                      Row currentRow = iterator.next();
                      Iterator<Cell> cellIterator = currentRow.iterator();
 
                          String keys="",values="";
                          int flags=1;
                          while (cellIterator.hasNext()) {
                    	  Cell currentCell = cellIterator.next();
                              if(flags==1) {
                                  keys=currentCell.getStringCellValue();
                              }else if(flags==3) {
                                  values=currentCell.getStringCellValue();
                              }
                              flags++;
                             maps.put(keys, values);
                      }
                  }    
                  
                  maps.remove("");
              } catch (FileNotFoundException e) {
                  e.printStackTrace();
              } catch (IOException e) {
                  e.printStackTrace();
              }
              return maps;
          }
          
          
          
}
