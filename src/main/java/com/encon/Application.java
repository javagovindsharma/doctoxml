package com.encon;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import com.encon.service.DocToXmlConverter;
import com.encon.xml.ReadXmlFile;

/**
 * To execute DOC to XML converter
 * 
 * @author govind.sharma
 */
public class Application {

	static final Logger logger = Logger.getLogger(Application.class);

	public static void main(String[] args) {
		BasicConfigurator.configure();
		try {

			if (args[0] != null) {

				String path = System.getProperty("user.dir");
				String fileName = args[0];

				String nameOfFile[] = fileName.split("\\.(?=[^\\.]+$)");

				if (nameOfFile[1].equals("docx"))
					new DocToXmlConverter().processDocxToXml(path, nameOfFile[0]);
				else if (nameOfFile[1].equals("doc"))
					new DocToXmlConverter().processDocToXml(path, nameOfFile[0]);
				else if (nameOfFile[1].equals("xml"))
					new ReadXmlFile().xmlToString(path, nameOfFile[0]);
				else
					throw new Exception("please provide Correct File Extension");
			} else {
				throw new Exception();
			}
		} catch (Exception e) {
			logger.error("Please Provide File Name " + e.getMessage());
			System.exit(0);
		}

	}

}
