package cager.parser.test;

import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
  
import javax.xml.parsers.DocumentBuilder;  
import javax.xml.parsers.DocumentBuilderFactory;  
import javax.xml.parsers.ParserConfigurationException;  
import javax.xml.xpath.XPathConstants;  
import javax.xml.xpath.XPathExpressionException;  
import javax.xml.xpath.XPathFactory;  
  
import org.jdom2.input.SAXBuilder;
import org.w3c.dom.Document;  
import org.w3c.dom.Element;  
import org.w3c.dom.Node;  
import org.w3c.dom.NodeList;  
import org.xml.sax.InputSource;  
import org.xml.sax.SAXException;  


public class XMLTester {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {  
		
			// Implementación DOM por defecto de Java  
	        // Construimos nuestro DocumentBuilder  
	        DocumentBuilder documentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();  
	        // Procesamos el fichero XML y obtenemos nuestro objeto Document  
	        Document doc = documentBuilder.parse(new InputSource(new FileInputStream("archivo.kdm")));  
	        //Document doc = documentBuilder.parse(new InputSource(new FileInputStream("E:\\WorkspaceParser\\VBParser\\ejemplosKDM\\archivo.kdm")));  
        
	        // Obtenemos la etiqueta raiz  
	         Element elementRaiz = doc.getDocumentElement();  
	         // Iteramos sobre sus hijos  
	         NodeList hijos = elementRaiz.getChildNodes();  
	         for(int i=0;i<hijos.getLength();i++){  
	            Node nodo = hijos.item(i);  
	            if (nodo instanceof Element){  
	               System.out.println(nodo.getNodeName());  
	            }  
	         }  
	           
	         // Buscamos una etiqueta dentro del XML  
	         NodeList listaNodos = doc.getElementsByTagName("codeElement");  
	         for(int i=0;i<listaNodos.getLength();i++){  
	            Node nodo = listaNodos.item(i);  
	            if (nodo instanceof Element){  
	               System.out.println(nodo.getTextContent());  
	            }  
	         }  
        
		   } catch (FileNotFoundException e) {  
	         e.printStackTrace();  
	       } catch (SAXException e) {  
	         e.printStackTrace();  
	       } catch (IOException e) {  
	         e.printStackTrace();  
	       } catch (ParserConfigurationException e) {  
	         e.printStackTrace();  
	       }   
	}

}
