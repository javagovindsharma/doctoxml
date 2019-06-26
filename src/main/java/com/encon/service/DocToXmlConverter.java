package com.encon.service;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.log4j.Logger;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * DOC to XML converter service
 * 
 * @author govind.sharma
 *
 */

/**
 * @author govind.sharma
 *
 */
public class DocToXmlConverter {

    static final Logger logger = Logger.getLogger(DocToXmlConverter.class);

    DocumentBuilderFactory docFactory = null;
    DocumentBuilder docBuilder = null;
    Element rootElement = null;
    Document docxml = null;
    boolean subHeaders = false;
    Element UrlElement = null;

    /**
     * @param path
     * @param fileName
     */
    public void processDocxToXml(String path, String fileName) {

	XWPFDocument xdoc = null;
	FileInputStream fis = null;
	List<XWPFTable> listatab = null;
	String fullPath = path + "/" + fileName + ".docx";

	try {
	    // Read file
	    fis = new FileInputStream(fullPath);
	    xdoc = new XWPFDocument(OPCPackage.open(fis));

	    initializeXml();
	    // get Document Body Paragraph content

	    Iterator<IBodyElement> iter = xdoc.getBodyElementsIterator();
	    while (iter.hasNext()) {

		String styleName = null, paraText = "", bulletsPoints = null, bulletText = "";
		boolean paragraphcIsBold = false;

		IBodyElement elem = iter.next();

		if (elem instanceof XWPFParagraph) {
		    styleName = ((XWPFParagraph) elem).getStyle();
		    paraText = ((XWPFParagraph) elem).getParagraphText();
		    bulletsPoints = ((XWPFParagraph) elem).getNumFmt();

		    if (bulletsPoints != null) {
			if (bulletsPoints.equalsIgnoreCase("bullet")) {
			    byte[] arr = ((XWPFParagraph) elem).getNumLevelText().getBytes();
			    bulletText = convertToHex(arr);
			}
		    }
		}

		createXmlTags(styleName, paraText, bulletsPoints, paragraphcIsBold, bulletText);

		if (listatab != null) {
		    listatab = null;
		}

	    }

	    // write the content into XML file
	    generateXml(path, fileName);
	    logger.info("Doc to Xml Convertion completed.");

	} catch (Exception ex) {
	    logger.error("Exception while generating XML from DOC" + ex.getMessage());
	    System.exit(0);
	}
    }

    /**
     * @param path
     * @param fileName
     */
    public void processDocToXml(String path, String fileName) {
	HWPFDocument doc = null;
	String fullPath = path + "/" + fileName + ".doc";

	WordExtractor we = null;
	try {
	    POIFSFileSystem fis = new POIFSFileSystem(new FileInputStream(fullPath));
	    doc = new HWPFDocument(fis);
	} catch (Exception e) {
	    logger.error("Unable to Read File..." + e.getMessage());
	    System.exit(0);
	}
	try {

	    we = new WordExtractor(doc);
	    Range range = doc.getRange();

	    initializeXml();

	    String[] paragraphs = we.getParagraphText();

	    for (int i = 0; i < paragraphs.length; i++) {
		org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);

		int j = 0;
		while (true) {
		    CharacterRun run = pr.getCharacterRun(j++);

		    StyleDescription style = doc.getStyleSheet().getStyleDescription(run.getStyleIndex());
		    String styleName = style.getName();
		    String paraText = run.text();
		    String bulletsPoints = null;

		    boolean paragraphcIsBold = run.isBold();
		    if (pr.getIlfo() != 0)
			bulletsPoints = "bullets";

		    createXmlTags(styleName, paraText, bulletsPoints, paragraphcIsBold, "");

		    if (run.getEndOffset() == pr.getEndOffset()) {
			break;
		    }
		}
	    }

	    generateXml(path, fileName);

	    logger.info("Document to Xml Convertion completed.");
	} catch (Exception ex) {
	    logger.error("Exception while generating XML from DOC" + ex.getMessage());
	    System.exit(0);
	}
    }

    /**
     * 
     */
    private void initializeXml() {

	// initialize XML Document
	try {
	    docFactory = DocumentBuilderFactory.newInstance();
	    docBuilder = docFactory.newDocumentBuilder();
	    docxml = docBuilder.newDocument();

	    rootElement = docxml.createElement("BENEFIT");
	    docxml.appendChild(rootElement);
	} catch (ParserConfigurationException e) {
	    logger.error("Exception while initializing XML" + e.getMessage());
	}

    }

    private void createXmlTags(String styleName, String paragraphText, String bulletsPoints, boolean paragraphcIsBold,
	    String bulletText) {

	// create XML Tags

	if (styleName != null && paragraphText.length() > 1) {

	    if (styleName != null && bulletsPoints != null) {
		appendListTags(paragraphText, bulletText);
	    } else if (paragraphcIsBold) {
		Element pragElement = docxml.createElement("B");
		pragElement.appendChild(docxml.createTextNode(paragraphText));

		rootElement.appendChild(pragElement);
		subHeaders = true;

	    } else if (styleName.equalsIgnoreCase("Style4")) {
		Element pragElement = docxml.createElement("TITLE");
		pragElement.appendChild(docxml.createTextNode(paragraphText.trim()));
		rootElement.appendChild(pragElement);
		subHeaders = true;

	    } else if (styleName.equalsIgnoreCase("Default")) {
		Element pragElement = docxml.createElement("P");
		pragElement.appendChild(docxml.createTextNode(paragraphText));

		rootElement.appendChild(pragElement);
		subHeaders = true;

	    } else if (styleName.equalsIgnoreCase("Normal")) {
		Element pragElement = docxml.createElement("P");
		pragElement.appendChild(docxml.createTextNode(paragraphText));

		rootElement.appendChild(pragElement);
		subHeaders = true;

	    } else if (styleName.equalsIgnoreCase("BodyCopy")) {
		Element pragElement = docxml.createElement("PS");
		pragElement.appendChild(docxml.createTextNode(paragraphText));
		rootElement.appendChild(pragElement);
		subHeaders = true;

	    } else if (styleName.equalsIgnoreCase("ListParagraph")) {
		appendListTags(paragraphText, bulletText);
	    } else if (styleName.equalsIgnoreCase("Subheader1")) {

		appendListTags(paragraphText, bulletText);

	    } else {
		Element pragElement = docxml.createElement("PS");
		pragElement.appendChild(docxml.createTextNode(paragraphText));

		rootElement.appendChild(pragElement);
		subHeaders = true;

	    }

	} else if (paragraphText.trim().length() > 1) {
	    Element pragElement = docxml.createElement("P");
	    pragElement.appendChild(docxml.createTextNode(paragraphText));

	    rootElement.appendChild(pragElement);
	    subHeaders = true;

	}

	if (subHeaders) {
	    Element pragElement = docxml.createElement("NEWLINE");
	    pragElement.appendChild(docxml.createTextNode(""));
	    rootElement.appendChild(pragElement);
	}
    }

    private void appendListTags(String paragraphText, String bulletText) {

	try {
	   
	    Element LiElement = docxml.createElement("LI");
	    LiElement.appendChild(docxml.createTextNode(paragraphText));

	    UrlElement = docxml.createElement("UL");
	    UrlElement.appendChild(LiElement);

	    rootElement.appendChild(UrlElement);
	    subHeaders = false;
	} catch (Exception e) {
	    e.printStackTrace();
	    System.exit(0);
	}

    }

    /**
     * @param path
     * @param fileName
     */
    private void generateXml(String path, String fileName) {
	try {
	    // write the content into xml file
	    TransformerFactory transformerFactory = TransformerFactory.newInstance();
	    Transformer transformer = transformerFactory.newTransformer();
	    transformer.setOutputProperty(OutputKeys.METHOD, "xml");
	    transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
	    transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
	    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
	    transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");

	    DOMSource source = new DOMSource(docxml);
	    StreamResult result = new StreamResult(new File(path + "/" + fileName + ".xml"));
	    transformer.transform(source, result);

	} catch (Exception e) {
	    logger.error("Exception while generating XML" + e.getMessage());
	}
    }

    @SuppressWarnings("static-access")
    public static String convertToHex(byte[] byteArray) {
	String hexString = "";

	for (int i = 0; i < byteArray.length; i++) {
	    String thisByte = "".format("%x", byteArray[i]);
	    hexString += thisByte;
	}

	return hexString;
    }

}
