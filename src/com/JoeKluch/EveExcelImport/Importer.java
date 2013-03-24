package com.JoeKluch.EveExcelImport;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

import java.net.MalformedURLException;
import java.net.URL;
import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File; 
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import jxl.*; 
import jxl.read.biff.BiffException;
import jxl.write.*; 
import jxl.write.Number;
 
public class Importer {
 
	public static void main(String argv[]) throws BiffException, IOException, SAXException, ParserConfigurationException{
 
		String Surl = "http://api.eve-central.com/api/marketstat?typeid=";
		String Eurl = "&usesystem=30000142";
		int Murl = 34;
		String markURL = "http://api.eve-marketdata.com/api/type_id.xml?char_name=Joe%20Kluch&v=";
		int i=0;
		
		try{
			//Set up the spreadsheets
			Workbook readWorkbook = Workbook.getWorkbook(new File("input.xls"));
			Sheet readSheet = readWorkbook.getSheet(0);
			//WritableWorkbook writeWorkbook = Workbook.createWorkbook(new File("output.xls"));
			WritableWorkbook writeWorkbook = Workbook.createWorkbook(new File("output.xls"), readWorkbook);
			//WritableSheet writeSheet = writeWorkbook.createSheet("First Sheet", 0);
			WritableSheet writeSheet = writeWorkbook.getSheet(0); 
			
			//Run for loop with try catch to take cell aN and dump the min sale in cN
			try{
				
				while(true){
					Cell a1 = readSheet.getCell(0,i);
					String stringAn = a1.getContents();
					System.out.println(stringAn);
					
					//Import Label
					//Label labelA = new Label(0, i, stringAn); 
					//writeSheet.addCell(labelA);
					if(!inBlackLst(stringAn)){
						//Get the typeID
						System.out.println("The item name is: "+markURL+stringAn.replace(" ", "+"));
						Murl = importXML((markURL+stringAn.replace(" ", "+")), "emd", "marketdata").intValue();
						System.out.println("The item ID is: "+Murl);
						//Import Val
						Number numberC = new Number(2, i, importXML((Surl+Murl+Eurl), "sell", "evecentral"));
						writeSheet.addCell(numberC);
					}
					//Number number = new Number(2, i, 3.1459); 
					//writeSheet.addCell(number);
					if(i==20){
						writeWorkbook.write(); 
						writeWorkbook.close();
						break;
					}
					i++;
				}
			}
			catch (Exception e) {
				writeWorkbook.write(); 
				writeWorkbook.close();
				System.err.println("End of full cells A"+i);
				System.err.println(e.getMessage());
			}
			/*
			Cell a1 = sheet.getCell(0,0);
			Cell b2 = sheet.getCell(1,1);
			//Cell c2 = sheet.getCell(2,1);
			
			String stringa1 = a1.getContents();
			String stringb2 = b2.getContents();
			System.out.println(stringa1);
			System.out.println(stringb2);
			//String stringb2 = b2.getContents();
			//String stringc2 = c2.getContents();
			
			if (a1.getType() == CellType.LABEL){
				LabelCell lc = (LabelCell) a1;
				stringa1 = lc.getString();
				System.out.println(stringa1);
			} 
			
			if (b2.getType() == CellType.NUMBER){
				NumberCell nc = (NumberCell) b2;
				numberb2 = nc.getValue();
			}*/
		}
		catch (Exception e) {
            System.err.println("Workbook fail");
            System.err.println(e.getMessage());
        }
		
		
		/*
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.parse(new URL(Surl+Murl+Eurl).openStream());
		doc.getDocumentElement().normalize();
		
		System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
		NodeList nList = doc.getElementsByTagName("sell");
		System.out.println("-----------------------");
		
		for (int temp = 0; temp < nList.getLength(); temp++) {
			
			Node nNode = nList.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				
				Element eElement = (Element) nNode;
				
				System.out.println("Average: " + getTagValue("avg", eElement));
				System.out.println("Min: " + getTagValue("min", eElement));
				System.out.println("Max: " + getTagValue("max", eElement));
				System.out.println("Median : " + getTagValue("median", eElement));
				
			}
		}
		*/
	}
	
private static boolean inBlackLst(String productName) {
	// TODO Auto-generated method stub
	try{
		// Open the file that is the first 
		// command line parameter
		FileInputStream fstream = new FileInputStream("blacklist.txt");
		// Get the object of DataInputStream
		DataInputStream in = new DataInputStream(fstream);
		BufferedReader br = new BufferedReader(new InputStreamReader(in));
		String strLine;
		//Read File Line By Line
		while ((strLine = br.readLine()) != null){
			// Print the content on the console
			if(productName.equalsIgnoreCase(strLine) || productName==""){
				return true;
			}
		}
		//Close the input stream
		in.close();
	}
	catch(Exception e){//Catch exception if any
		System.err.println("Error: " + e.getMessage());
	}
	return false;
}

private static String getTagValue(String sTag, Element eElement) {
	NodeList nlList = eElement.getElementsByTagName(sTag).item(0).getChildNodes();
	
	Node nValue = (Node) nlList.item(0);
	
	return nValue.getNodeValue();
}

private static Double importXML(String site, String tag, String theSite) throws ParserConfigurationException, MalformedURLException, SAXException, IOException{
	
	if(theSite=="evecentral"){
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.parse(new URL(site).openStream());
		doc.getDocumentElement().normalize();
		
		System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
		NodeList nList = doc.getElementsByTagName(tag);
		System.out.println("-----------------------");
		
		for (int temp = 0; temp < nList.getLength(); temp++) {
			
			Node nNode = nList.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				
				Element eElement = (Element) nNode;
				
				System.out.println("Average: " + getTagValue("avg", eElement));
				System.out.println("Min: " + getTagValue("min", eElement));
				System.out.println("Max: " + getTagValue("max", eElement));
				System.out.println("Median : " + getTagValue("median", eElement));
				return Double.parseDouble(getTagValue("min", eElement));
			}
		}
	}
	else if(theSite=="marketdata"){
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.parse(new URL(site).openStream());
		doc.getDocumentElement().normalize();
		
		System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
		NodeList nList = doc.getElementsByTagName(tag);
		System.out.println("-----------------------");
		
		for (int temp = 0; temp < nList.getLength(); temp++) {
			
			Node nNode = nList.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				
				Element eElement = (Element) nNode;
				
				System.out.println("Value : " + getTagValue("val", eElement));
				return Double.parseDouble(getTagValue("val", eElement));
			}
		}
	}
	return 404.0;

}


 
}