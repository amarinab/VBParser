/*
 * Copyright 2005-2006 Paul Cager.
 * 
 * www.paulcager.org
 * 
 * This file is part of cager.parser.
 * 
 * cager.parser is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 * 
 * cager.parser is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with cager.parser; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
*/
/**
 * 
 */
package cager.parser.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.JDOMException;
import org.jdom2.Namespace;
import org.jdom2.input.SAXBuilder;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

import junit.framework.Assert;
import cager.parser.SimpleNode;
import cager.parser.VBErrHandlerVisitor;
import cager.parser.VBParser;
import cager.parser.VBParserVisitor;

/**
 * @author paul
 *
 */
public class SimpleTest
{

    private void simpleTest() throws Exception
    {
        VBParser parser;
        InputStream is = getClass().getResourceAsStream("testvb.cls");
        
        parser = new VBParser(is);
        parser.setAttemptErrorRecovery(true);
        SimpleNode node = parser.CompilationUnit(false);
        
        XMLParser(node);
        
        //node = (SimpleNode)(parser.ExprTest());
        //          PrintWriter ostr = new PrintWriter(new FileWriter(args[1]));

        //        node.jjtAccept(new VBErrHandlerVisitor(), null);

        
        //VBParserVisitor v = new VBErrHandlerVisitor();
        //node.jjtAccept(v, null);
        
//        node.dump(">");
//
        //System.out.println(node.allText(true));
        //System.out.println("VBParser:  Parsing completed successfully.");

        System.out.println("Completed test of generated Parser");
    }
    
    public static void main(String[] args) throws Exception
    {
        new SimpleTest().simpleTest();
        System.out.println("Finished");
    }
    
    
    public void XMLParser (SimpleNode astNode) {

    	// Creamos el builder basado en SAX  
    	SAXBuilder builder = new SAXBuilder();
    	
    	Element segmento = inicializaKDM("testvb");

    	
    	trataNodo(astNode,segmento);
    	
    	
    	System.out.println(elementoToString(segmento));
		
    }

    
    
    private void trataNodo(SimpleNode astNode, Element segmento) {
		
    	int numChildren = astNode.jjtGetNumChildren();
    	
    	for (int i=0;i<numChildren; i++ ){
    		
    		System.out.println(astNode.jjtGetChild(i).getClass().toString());
    		
    		trataNodo((SimpleNode)astNode.jjtGetChild(i),segmento);

/*  		switch (){
			
    		case value:
				
				break;

			default:
				break;
			}
 */ 		
     	}
	}

	private String elementoToString(Element segmento) {
    	
    	Document documentJDOM = new Document();
		documentJDOM.addContent(segmento);
		// Vamos a serializar el XML  
		// Lo primero es obtener el formato de salida  
		// Partimos del "Formato bonito", aunque también existe el plano y el compacto  
		Format format = Format.getPrettyFormat();  
		// Creamos el serializador con el formato deseado  
		XMLOutputter xmloutputter = new XMLOutputter(format);  
		// Serializamos el document parseado  
		String docStr = xmloutputter.outputString(documentJDOM);
		
		return docStr;
	}

    
	private Element inicializaKDM(String segmentName){
    	
    	// Creamos el builder basado en SAX
    	SAXBuilder builder = new SAXBuilder();  
    	
    	Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
		Namespace xsi = Namespace.getNamespace("xsi","http://www.w3.org/2001/XMLSchema-instance");
		Namespace action = Namespace.getNamespace("action","http://kdm.omg.org/action");
		Namespace code = Namespace.getNamespace("code","http://kdm.omg.org/code");
		Namespace kdm = Namespace.getNamespace("kdm","http://kdm.omg.org/kdm");
		Namespace source = Namespace.getNamespace("source","http://kdm.omg.org/source");
		
		Element segmento = new Element("Segment",kdm);  
		segmento.addNamespaceDeclaration(xmi);
		segmento.addNamespaceDeclaration(xsi);
		segmento.addNamespaceDeclaration(action);
		segmento.addNamespaceDeclaration(code);
		segmento.addNamespaceDeclaration(kdm);
		segmento.addNamespaceDeclaration(source);
		//segmento.setName("HelloWorld_Example");
		segmento.setAttribute("name", segmentName);

		Element modelo = new Element ("model");
		modelo.setAttribute("id", "id.0", xmi);
		modelo.setAttribute("name", segmentName);
		modelo.setAttribute("type", "code:CodeModel", xmi);
		
		
		segmento.addContent(modelo);
		
		return segmento;
    }
    
}

