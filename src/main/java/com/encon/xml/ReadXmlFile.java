package com.encon.xml;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Map;
import org.apache.log4j.Logger;

/**
 * @author govind.sharma
 *
 */
public class ReadXmlFile {
    static final Logger logger = Logger.getLogger(ReadXmlFile.class);

    public Map<String, String> maps = new HashMap<String, String>();

    public ReadXmlFile() {
	maps = new ExcelReaderFile().readExcelFile();
   }

    public void xmlToString(String path, String fileName) {
	try {
	    long start= System.currentTimeMillis() ;
	    logger.info("Reading XML File ..");
	    String strContent=  readingXMlFileConvertString(path, fileName);
	    logger.info("Generating  XML File ..");
	    stringToXmlFileGerator(path,fileName,strContent);
	    long end= System.currentTimeMillis() ;
	    logger.info("Total Processing Time Occured :"+(end-start));
	} catch (Exception e) {
	    logger.error("Exception while generating XML from DOC" + e.getMessage());
	    System.exit(0);
	}
    }

    private String readingXMlFileConvertString(String path, String fileName) {
	String xml2String = "", resultXml = "";
	try {
	    // our XML file for this example
	    File xmlFile = new File(path + "/" + fileName + ".xml");
	    // Let's get XML file as String using BufferedReader FileReader uses platform's
	    // default character encoding if you need to specify a different encoding, use
	    InputStreamReader fileReader = new FileReader(xmlFile);
	    BufferedReader bufReader = new BufferedReader(fileReader);
	    StringBuilder sb = new StringBuilder();
	    String line = bufReader.readLine();
	    while (line != null) {
		sb.append(line).append("\n");
		line = bufReader.readLine();

	    }
	    xml2String = sb.toString();

	    for (int i = 0; i < xml2String.length(); i++) {
		String ch = String.valueOf(xml2String.charAt(i));
		// using for-each loop for iteration over Map.entrySet()
		for (Map.Entry<String, String> entry : maps.entrySet()) {
		    if (ch.contains(entry.getKey())) {
			
			System.err.println(entry.getKey()+"--->>"+entry.getValue());
                		if(ch.contains("’")){
                		    System.err.println("char val "+ch+"  "+ch.hashCode());
                		    
                		}
			ch = entry.getValue();
		    }
		}
		resultXml += ch;
	    }
	    bufReader.close();

	} catch (Exception e) {
	    logger.error("Exception while generating XML from DOC" + e.getMessage());
	      System.exit(0);
	}
	return resultXml;
    }

    
    private  void stringToXmlFileGerator(String path, String fileName,String strContent) {
	      
		logger.info(getClass().getName()+" stringToXmlFileGerator  method  execution ");
	      try {
	           FileOutputStream out = new FileOutputStream(path+"/new"+fileName+".xml");
                    out.write(strContent.getBytes());
                    out.close();
                    logger.info("File Completly Generated");
            }catch (Exception e) {
        	   logger.error("Exception while generating String  from XML" + e.getMessage());
        	   e.printStackTrace();
     	          System.exit(0);
	    }
    }   
    
}
