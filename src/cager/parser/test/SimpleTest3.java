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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.jdom2.Attribute;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.JDOMException;
import org.jdom2.Namespace;
import org.jdom2.filter.Filters;
import org.jdom2.input.SAXBuilder;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;
import org.jdom2.xpath.XPathExpression;
import org.jdom2.xpath.XPathFactory;

import cager.parser.ASTAssignment;
import cager.parser.ASTBinOp;
import cager.parser.ASTDeclaration;
import cager.parser.ASTDeclare;
import cager.parser.ASTExpression;
import cager.parser.ASTFormalParamList;
import cager.parser.ASTIfStatement;
import cager.parser.ASTName;
import cager.parser.ASTParamSpec;
import cager.parser.ASTPrimaryExpression;
import cager.parser.ASTProcDeclaration;
import cager.parser.ASTStatement;
import cager.parser.ASTTypeName;
import cager.parser.ASTVarDecl;
import cager.parser.SimpleNode;
import cager.parser.VBParser;


/**
 * @author paul
 *
 */
public class SimpleTest3 
{
	//TODO: Implementar estructura de control de índices de Identificadores.
	private int idCount = 0;
	private int idName = 0;
	private int idAction = 0;
	
	Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
	String contextoLocal = "";
	Element segmento;
	
//	static Logger logger = LoggerFactory.getLogger(SimpleTest2.class);
	
	public HashMap<String, Integer> hsmMethodUnits = new HashMap<String, Integer>(); 
	public HashMap<Integer,Element> hsmMethodUnitsObj = new HashMap<Integer,Element>();  
	
	public HashMap<String, Integer> hsmStorableUnits = new HashMap<String, Integer>(); 
	public HashMap<Integer,Element> hsmStorableUnitsObj = new HashMap<Integer,Element>();
	public HashMap<String, Integer> hsmLanguajeUnitDataType = new HashMap<String, Integer>(); 
	
	public Element elementoTipos = null;
	
	//Tipos de Visual Basic
	//"Boolean","Byte","Char","Date","Decimal",
	//   "Double","Integer","Long","Short","Single","Object","String"
	String dataTypes[] = {"Boolean","Byte","Char","Date","Decimal",
			   			  "Double","Integer","Long","Short","String","Void"};//,"Object"};
	
	
	public String archivo = "testvbPRU.cls";
	//public String archivo = "testvb.cls";
	//public String archivo = "main.frm";
	
	// Pc Fijo
	//public String ruta = "E:\\WorkspaceParser\\";
	// Pc Portatil
	public String ruta = "C:\\Users\\Tony\\Dropbox\\WorkspaceParser\\";
	
	private  boolean debuggMode = false;

		
    private void simpleTest() throws Exception
    {
        VBParser parser;
      
        InputStream is = getClass().getResourceAsStream(archivo);

        
        // Genera el nodo desde el código Visual Basic
        parser = new VBParser(is);
        parser.setAttemptErrorRecovery(true);
        SimpleNode node = parser.CompilationUnit(false);
        
        XMLParser(node);
        
        XMLtoJavaParser();
        
       // logger.info("Completed test of generated Parser");
        System.out.println("Completed test of generated Parser");
        
    }

	public static void main(String[] args) throws Exception
    {
        new SimpleTest3().simpleTest();
        //logger.info("Finished");
        System.out.println("Finished"); 
    }
    
    public void XMLParser (SimpleNode astNode) {

    	// Creamos el builder basado en SAX  
    	SAXBuilder builder = new SAXBuilder();
    	 
     	Element segmento = inicializaKDM(archivo);

        	
    	trataNodo(astNode,segmento,0,"");
    	
    	trataNodo2(astNode,segmento,0,0,"");
    	
    	//logger.info(elementoToString(segmento));
    	System.out.println(elementoToString(segmento));
		
    }

    private void trataNodo(SimpleNode astNode, Element segmento, int nivel, String contexto) {
    	
    	int numChildren = astNode.jjtGetNumChildren();
    	//TODO Solucionar referencias a Namespaces 
    	Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
    	
    	String contextoLocal = "";
    	contextoLocal = contexto;
    	
    	String strNombre = "";
    	
    	for (int i=0;i<numChildren; i++ ){
    		
    		strNombre = ((SimpleNode)astNode.jjtGetChild(i)).getName();
    		
    		if(astNode.jjtGetChild(i).getClass() == ASTProcDeclaration.class ){ 

  			
for (int j=0; j<nivel; j++) System.out.print("\t");
//logger.info("1. ->"+astNode.jjtGetChild(i).getClass().toString()+" - "+strNombre+" - "+" Nivel: "+nivel);
System.out.println("1. ->"+astNode.jjtGetChild(i).getClass().toString()+" - "+strNombre+" - "+" Nivel: "+nivel);
    			
    			Element codeElement = new Element ("codeElement");
    			codeElement.setAttribute("id", "id."+idCount, xmi);
    			hsmMethodUnits.put(strNombre, idCount);
    			hsmMethodUnitsObj.put(idCount, codeElement);
    			contextoLocal = strNombre;
    			idCount++;
    			codeElement.setAttribute("name", strNombre);
    			codeElement.setAttribute("type", "code:MethodUnit", xmi);    			
    			//TODO: Comprobar que tipo de method es: Enumeration MethodKind
    			codeElement.setAttribute("kind", "method");
    			
    			ASTProcDeclaration astProcDecl = (ASTProcDeclaration)astNode.jjtGetChild(i);
    			String contenido = astProcDecl.allText(true);
    			
    			Element sourceElement = new Element ("source");
    			sourceElement.setAttribute("id", "id."+idCount, xmi);
    			idCount++;
    			sourceElement.setAttribute("language", "VisualBasic");
    			sourceElement.setAttribute("snippet", contenido);
    			
    			
    			codeElement.addContent(sourceElement);
    			
    			segmento.getChild("model").getChild("codeElement").addContent(codeElement);
    			this.segmento = segmento;
    			
//<source xmi:id="id.41" language="Hla" snippet="Field5[1] -> Acc_Type[0]"/>
    			
    			//Buscar Parámetros para signature
    			
    			codeElement.addContent(TratarSignature((SimpleNode)astNode.jjtGetChild(i),strNombre));
    			
    		}else if(astNode.jjtGetChild(i) instanceof ASTName){

    			
    			if(astNode instanceof ASTDeclare
    					&& astNode.jjtGetChild(i).jjtGetParent().jjtGetChild(0).equals(astNode.jjtGetChild(i)) ){ //TODO: Revisar si esto es lógico.
    				
    				ASTDeclare astDeclare = (ASTDeclare) astNode;
    				Element codeElement = new Element ("codeElement");
    				codeElement.setAttribute("id", "id."+idCount, xmi); 
    				hsmMethodUnits.put(strNombre, idCount);
    				hsmMethodUnitsObj.put(idCount, codeElement);
    				contextoLocal = strNombre;			
    				idCount++;
    				codeElement.setAttribute("name", strNombre);
    				codeElement.setAttribute("export",astDeclare.getFirstToken().toString().toLowerCase());
    				codeElement.setAttribute("type", "code:MethodUnit", xmi);    			
    				//TODO: Comprobar que tipo de method es: Enumeration MethodKind
        			codeElement.setAttribute("kind", "method"); 
    				
        			
        			String contenido = astDeclare.allText(true);
        			
        			Element sourceElement = new Element ("source");
        			sourceElement.setAttribute("id", "id."+idCount, xmi);
        			idCount++;
        			sourceElement.setAttribute("language", "VisualBasic");
        			sourceElement.setAttribute("snippet", contenido);
        			
        			
        			codeElement.addContent(sourceElement);
        			
        			
        			
    				segmento.getChild("model").getChild("codeElement").addContent(codeElement);
    				this.segmento = segmento;
    				
    			
    				//Buscar Parámetros para signature
    				
    				codeElement.addContent(TratarSignature(astNode,strNombre));
    				
    			}else if(astNode instanceof ASTVarDecl){
    				
    				ASTDeclaration astDeclaration = null;
    				String exportVal = null;
    				if(astNode.jjtGetParent() instanceof ASTDeclaration){
    					astDeclaration = (ASTDeclaration) astNode.jjtGetParent(); 
    					exportVal = astDeclaration.getFirstToken().toString().toLowerCase();
    				}
    				
    				Element codeElement = new Element ("codeElement");
        			codeElement.setAttribute("id", "id."+idCount, xmi);
        			String tipo = null;
        			
        		    //Buscamos el tipo de la variable declarada recorriendo todos los hijos buscando el ASTTypeName
    				for(int n = 0; n<astNode.jjtGetNumChildren() && tipo==null;n++){
    					if(astNode.jjtGetChild(n) instanceof ASTTypeName){
    						tipo = ((SimpleNode)((SimpleNode) astNode.jjtGetChild(n)).jjtGetChild(0)).getName();
    					}
    				}
    				
    				//Si no se ha encontrado comprobamos si es una declaración múltiple y buscamos el tipo
    				if(tipo==null){
    					
    					SimpleNode Padre = (SimpleNode) astNode.jjtGetParent();
    					
    					for(int t = 0; t<Padre.jjtGetNumChildren() && tipo==null;t++){
   						
    						SimpleNode Hijo = (SimpleNode) Padre.jjtGetChild(t);
    						
    						for(int n = 0; n<Hijo.jjtGetNumChildren();n++){
    							
    							if(Hijo.jjtGetChild(n) instanceof ASTTypeName){
    	    						tipo = ((SimpleNode)((SimpleNode) Hijo.jjtGetChild(n)).jjtGetChild(0)).getName();
    	    					}
    	    				
    						}
    						
    					}
    					
    				}

        			hsmStorableUnits.put(strNombre, idCount);
        			hsmStorableUnitsObj.put(idCount, codeElement);
        			        			
        			Integer numeroTipo = hsmLanguajeUnitDataType.get(tipo);
        			if(numeroTipo == null){ numeroTipo = 13;} 
        			
        			/*   				
    				if(numeroTipo == null){
    			
    					numeroTipo = idCount;
    					Element enumerated = new Element("codeElement");
    					enumerated.setAttribute("id", "id."+numeroTipo, xmi);
    					//enumerated.setAttribute("type", "code:ClassUnit", xmi);
    					enumerated.setAttribute("type", "code:StringType", xmi);
    					enumerated.setAttribute("name", tipo);
    					hsmLanguajeUnitDataType.put(tipo, numeroTipo);
    					idCount++;
    				
    					    					
    					//Añadimos el enumerated a la lista de tipos
    					elementoTipos.addContent(enumerated);
    					
    				}
        			
*/      			
					idCount++;
					codeElement.setAttribute("name", strNombre);
					//StorableUnit no permite atributo export
					//if(exportVal != null) codeElement.setAttribute("export", exportVal);
					codeElement.setAttribute("type", "code:StorableUnit", xmi);
					codeElement.setAttribute("type","id."+numeroTipo);

					if (contextoLocal == ""){
						codeElement.setAttribute("kind", "global");
						segmento.getChild("model").getChild("codeElement").addContent(codeElement);
						this.segmento = segmento;
					}else{
						codeElement.setAttribute("kind", "local");
						Element elementAux = hsmMethodUnitsObj.get(hsmMethodUnits.get(contextoLocal));//.getChild("codeElement");
						elementAux.addContent(codeElement);						
					}
    			}
    		}
    	
			trataNodo((SimpleNode)astNode.jjtGetChild(i),segmento,nivel+1,contextoLocal);

    	}
    }

    private Element TratarSignature(SimpleNode astNode, String strNombre) {
		
    	//Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
    	
    	Element signature = new Element("codeElement");
		signature.setAttribute("id", "id."+idCount, xmi);
		signature.setAttribute("type", "code:Signature", xmi);
		signature.setAttribute("name",strNombre);
		idCount++;
    	
    	for(int i = 0; i<astNode.jjtGetNumChildren();i++){
    		
    		if(astNode.jjtGetChild(i) instanceof ASTFormalParamList){
    			
    			SimpleNode paramList = (SimpleNode) astNode.jjtGetChild(i);
    			
    			for(int j=0;j<paramList.jjtGetNumChildren();j++){
    				
    				String tipo = "";
    	    		String nombre = "";
    	    		String valorTipo = "";
    	    		String kind = null;
    	    		
    				for(int k = 0; k < paramList.jjtGetChild(j).jjtGetNumChildren(); k++){
    					
    					SimpleNode hijo = (SimpleNode) paramList.jjtGetChild(j).jjtGetChild(k);
    					//TODO: hacer la correspondencia token/valor del ParameterKind (byVal/byValue, etc...)
    					
    					if(paramList.jjtGetChild(j) instanceof ASTParamSpec){
    						ASTParamSpec param = ((ASTParamSpec) paramList.jjtGetChild(j));
    					
	    					if(param.getByVal()){
	    						kind = "byValue";
	    					}else if(param.getbyRef()){
	    						kind = "byReference";
	    					}else{
	    						kind = null;
	    					}
    					}
    					//Se busca el nombre 
    					if(hijo instanceof ASTName){
							nombre = hijo.getName();
    					}else if(hijo instanceof ASTTypeName){
    							
    							if(hijo.jjtGetNumChildren()>1){
    								tipo =((SimpleNode) hijo.jjtGetChild(0)).getName();
    								for(int l=1;l<hijo.jjtGetNumChildren();l++){
    									tipo = tipo+"."+((SimpleNode) hijo.jjtGetChild(l)).getName();
    								}

    							}else{
    								tipo =((SimpleNode) hijo.jjtGetChild(0)).getName();
    							}
        				}else{
        					if(hijo instanceof ASTExpression){
        						valorTipo = ((SimpleNode) hijo.jjtGetChild(0)).getName();
        					}
    					}
    					
    				}
    			
    				Integer numeroTipo = hsmLanguajeUnitDataType.get(tipo);
    				
    				if(numeroTipo == null){
    			
    					numeroTipo = idCount;
    					Element enumerated = new Element("codeElement");
    					enumerated.setAttribute("id", "id."+numeroTipo, xmi);
    					enumerated.setAttribute("type", "code:EnumeratedType", xmi);
    					enumerated.setAttribute("name", tipo);
    					hsmLanguajeUnitDataType.put(tipo, numeroTipo);
    					idCount++;
    				
    					
    					Element valor = new Element("value");
    					valor.setAttribute("id", "id."+idCount, xmi);
    					valor.setAttribute("name", valorTipo);
    					valor.setAttribute("type","id."+String.valueOf(numeroTipo));
    					idCount++;
    						
    					enumerated.addContent(valor);

    					//Añadimos el enumerated a la lista de tipos
    					elementoTipos.addContent(enumerated);
    					
    				}
    				
    				Element parameterUnit = new Element("parameterUnit");
    				parameterUnit.setAttribute("id", "id."+idCount, xmi);
    				parameterUnit.setAttribute("name", nombre);
    				if(kind != null) parameterUnit.setAttribute("kind", kind);
    				parameterUnit.setAttribute("type", "id."+numeroTipo);
    				parameterUnit.setAttribute("pos", String.valueOf(j));
    				idCount++;
    				
    				signature.addContent(parameterUnit);
   				
    			}
    			
    			//Añadimos el tipo devuelto (tipo del Return)
    			String tipo = "Void";
    			for (int n = 1; n < astNode.jjtGetNumChildren(); n++){
    				
    				if(astNode.jjtGetChild(n) instanceof ASTName){
    					tipo = ((SimpleNode) astNode.jjtGetChild(n)).getName();
        			}
    				
    			}
    							
				Element parameterUnit = new Element("parameterUnit");
				parameterUnit.setAttribute("id", "id."+idCount, xmi);
				parameterUnit.setAttribute("type", "id."+hsmLanguajeUnitDataType.get(tipo));
				parameterUnit.setAttribute("kind", "return");
				idCount++;
				signature.addContent(parameterUnit);
    		}
    	}
    	return signature;
    }

	private void trataNodo2(SimpleNode astNode, Element segmento, int nivel, int nivelAcumulado, String contexto) {
		
    	int numChildren = astNode.jjtGetNumChildren();
    	//TODO Solucionar referencias a Namespaces 
    	
    	SimpleNode snHijo = new SimpleNode(0);
    	
    	String strNombre = "";
    	
    	contextoLocal = contexto;

    	strNombre = astNode.getName();
    	
    	if(hsmMethodUnits.get(strNombre)!=null){
    		contextoLocal=strNombre;    		
    	}

String espacios = "";
for(int j=0;j<nivelAcumulado;j++) espacios+="    ";
//logger.info(espacios+"Trata nodo : "+nivel+" - "+numChildren+" -> "+astNode.getClass().getName()+" "+astNode.toString());    	
//System.out.println(espacios+"Trata nodo : "+nivel+" - "+numChildren+" -> "+astNode.getClass().getName()+" "+astNode.toString());    	
System.out.println(espacios+nivel+" - "+numChildren+" -> "+astNode.getClass().getName()+" "+astNode.toString());    	

		// Recorre todos los hijos del ASTNode
    	for (int i=0;i<numChildren; i++ ){
    		
    		strNombre = ((SimpleNode)astNode.jjtGetChild(i)).getName();
  			
    		
    		if ((SimpleNode)astNode.jjtGetChild(i) instanceof ASTAssignment){
    			
    			ASTAssignment assigment = (ASTAssignment)astNode.jjtGetChild(i);
    			
    			
    			//Preparamos el ActionElement
    			Element actionElement = new Element ("codeElement");
                idAction = idCount;
                idCount++;
                actionElement.setAttribute("id", "id."+idAction, xmi); 
                
                
                trataExpressions(assigment, actionElement, nivelAcumulado);
    		
    		}else if((SimpleNode)astNode.jjtGetChild(i) instanceof ASTExpression){
                
    			ASTExpression expression = (ASTExpression)astNode.jjtGetChild(i);
    			
    			
    			//Preparamos el ActionElement
    			Element actionElement = new Element ("codeElement");
                idAction = idCount;
                idCount++;
                actionElement.setAttribute("id", "id."+idAction, xmi); 
                
                
                trataExpressions(expression, actionElement, nivelAcumulado);
    			
    		
    		}else if((SimpleNode)astNode.jjtGetChild(i) instanceof ASTStatement){
    			
    			ASTStatement astStatement = (ASTStatement) astNode.jjtGetChild(i);
    			//Preparamos el ActionElement
    			Element actionElement = new Element ("codeElement");
    			idCount++;
	            idAction = idCount;
	            idCount++;
	            actionElement.setAttribute("id", "id."+idAction, xmi); 
	            actionElement.setAttribute("type", "action:ActionElement", xmi);
				actionElement.setAttribute("name","a"+idName);
				//actionElement.setAttribute("name",contexto+"_"+idName);
				idName++;
    			
    			trataStatements(astStatement, actionElement);
    			
    		}
    		trataNodo2((SimpleNode)astNode.jjtGetChild(i),segmento,nivel+1,nivel+nivelAcumulado,contextoLocal);
     	}
	}
    
	private void trataExpressions (SimpleNode astAssignment, Element actionElement,int nivelAcumulado){
		
		SimpleNode snHijo = new SimpleNode(0);
		String strNombre = "";
		
		
		actionElement.setAttribute("type", "action:ActionElement", xmi);
        
		if(this.debuggMode)actionElement.setAttribute("debug",astAssignment.toString());                
		actionElement.setAttribute("name","a"+idName);
		idName++;
String espacios = "";
for(int j=0;j<nivelAcumulado;j++) espacios+="    ";
						
			//Recorre todos los nodos del ASTAssignment 
			for (int k=0;k<astAssignment.jjtGetNumChildren(); k++ ){
			
				 snHijo = (SimpleNode) astAssignment.jjtGetChild(k);
			
				//Tratamos el hijo ASTPrimaryExpresion
				if(snHijo instanceof ASTPrimaryExpression){
					
					SimpleNode snNieto;
					
					//Recorre todos los hijos del ASTPrimaryExpression
					for (int l=0;l<snHijo.jjtGetNumChildren(); l++ ){
						
						snNieto = (SimpleNode)snHijo.jjtGetChild(l);
						String nombreCompleto = "";
						
						for(int m = 0; m < snHijo.jjtGetNumChildren(); m++){
							if(m>0) nombreCompleto += ".";
							nombreCompleto += ((SimpleNode) snHijo.jjtGetChild(m)).getName(); 
						}
						
						//Comprobamos si se trata de un StorableUnit reconocido
						if(hsmStorableUnits.get(nombreCompleto) != null){
							
							
							
							//<actionRelation xmi:id="id.46" xmi:type="action:Writes" to="id.39"/>
							Element actionRelation = new Element ("actionRelation");
		        			actionRelation.setAttribute("id", "id."+idCount, xmi);
		        			idCount++;
		        			actionRelation.setAttribute("type", "action:Writes", xmi);
		        			actionRelation.setAttribute("to", "id."+String.valueOf(hsmStorableUnits.get(nombreCompleto)));
		         			actionRelation.setAttribute("from","id."+idAction);
		        						 
		         			actionElement.addContent(actionRelation);
		         			
		         			//TODO: Revisar qué hacemos con las referencias a tributos de Objetos (Objeto.atributo = ...) ASTName y ASTName
		         			//salir 
							//TODO: Tomamos solo uno de los ASTNames hasta que sepamos cómo hacer con las properties de un objeto.
							break;     			
						
						}else{
							
	//logger.info("Detectado un StorabeUnit nuevo!!!! ¿Qué hacemos?");
	System.out.println("Detectado un StorabeUnit nuevo!!!! ¿Que hacemos?: "+snNieto.getName());
	
	
							//Optamos temporalmente por añadirlo como StorableUnit global.
							//TODO: Revisar qué hacer en estos casos, ya que el parser no trata los "Begin"
	
							//Comprobamos si a
							Element codeElement = new Element ("codeElement");
							codeElement.setAttribute("id", "id."+idCount, xmi);
							
							hsmStorableUnits.put(nombreCompleto, idCount);
							hsmStorableUnitsObj.put(idCount, codeElement);
							
							idCount++;
							codeElement.setAttribute("name", nombreCompleto);
							codeElement.setAttribute("type", "code:StorableUnit", xmi);
							codeElement.setAttribute("type","");
								
							//codeElement.setAttribute("kind", "external");//TODO: Ver cómo tratar éste caso, variables utilizadas no definidas.
							//Para MethodUnit no hay external
							codeElement.setAttribute("kind", "unknown");
							
							segmento.getChild("model").getChild("codeElement").addContent(codeElement);
							this.segmento = segmento;
							
							
							
							//Añadimos el Action:Writes
							//TODO: Qué hacemos con los properties de un objeto.
							Element actionRelation = new Element ("actionRelation");
		        			actionRelation.setAttribute("id", "id."+idCount, xmi);
		        			idCount++;
		        			actionRelation.setAttribute("type", "action:Writes", xmi);
		        			actionRelation.setAttribute("to", "id."+String.valueOf(hsmStorableUnits.get(nombreCompleto)));
		         			actionRelation.setAttribute("from","id."+idAction);
		        						 
		         			actionElement.addContent(actionRelation);
							
							//salir 
							//TODO: Tomamos solo uno de los ASTNames hasta que sepamos cómo hacer con las properties de un objeto.
							break;
							//l=snHijo.jjtGetNumChildren();
	
						}
						
					}//FIN hijos ASTPrimaryExpression
				
					
				//Tratamos el caso de ASTExpression
				}else if(snHijo instanceof ASTExpression){
					
					SimpleNode snNieto;
					
					//Recorre todos los hijos del ASTExpression
					for (int l=0;l<snHijo.jjtGetNumChildren(); l++ ){
						
						snNieto = (SimpleNode)snHijo.jjtGetChild(l);
						
						//Si el hijo es un ASTName
						if(snNieto instanceof ASTName){
							
							//Comprobamos si es un storableUnit reconocido
							if(hsmStorableUnits.get(snNieto.getName()) != null){
								
								Element actionRelation = new Element ("actionRelation");
			        			actionRelation.setAttribute("id", "id."+idCount, xmi);
			        			idCount++;
			        			actionRelation.setAttribute("type", "action:Reads", xmi);
			        			actionRelation.setAttribute("to", "id."+String.valueOf(hsmStorableUnits.get(snNieto.getName())));
			         			actionRelation.setAttribute("from","id."+idAction);
			        						 
			         			actionElement.addContent(actionRelation);
			         			
			         			//TODO: Revisar qué hacemos con las referencias atributos de Objetos (Objeto.atributo = ...) ASTName y ASTName
							
			         		//Si no es un storableUnit reconocido
							}else{
	
								//logger.info("Encontrado un Hijo de ASTExpression que no es ASTName");
								System.out.println("Encontrado un Hijo de ASTExpression que no es ASTName");
							
							}//FIN Comprobar si se conoce el storableUnit
						
						//Si el hijo es ASTBinOp
						}else if(snNieto instanceof ASTBinOp){
							
							//Si el hijo es ASTName 
							if(snNieto.jjtGetChild(0) instanceof ASTName){
								
								strNombre= ((SimpleNode)(snNieto).jjtGetChild(0)).getName();
								
								//Tratamos la llamada a método
								
								//Si la MethodUnit no estaba registrada
								if(hsmMethodUnits.get(strNombre)==null && strNombre != null){
	//idCount++;	    			  					
									//Se da de alta como nueva MethodUnit
				  					Element nuevaFuncion = new Element ("codeElement");
				    				nuevaFuncion.setAttribute("id", "id."+idCount, xmi);
				    				hsmMethodUnits.put(strNombre, idCount);
				    				hsmMethodUnitsObj.put(idCount, nuevaFuncion);
	
	//logger.info("METODO: "+strNombre+": "+idCount);  				
	System.out.println("METODO: "+strNombre+": "+idCount);
	
									idCount++;
				    				nuevaFuncion.setAttribute("name", ""+strNombre);
				    				nuevaFuncion.setAttribute("type", "code:MethodUnit", xmi);    			
				    				//nuevaFuncion.setAttribute("kind", "system");  //TODO: Ver cómo tratar éste caso, funciones llamadas no definidas.
				    				//nuevaFuncion.setAttribute("kind", "unknown"); //TODO: Ver cómo tratar éste caso, funciones llamadas no definidas.
				    				//nuevaFuncion.setAttribute("kind", "external");  //Tratamos los métodos no conocidos como de tipo exteral.
				    				//Para los MethodUnit no hay external
				    				nuevaFuncion.setAttribute("kind", "unknown");
				    				
				    				segmento.getChild("model").getChild("codeElement").addContent(nuevaFuncion);
				    				this.segmento = segmento;
				  					
				    				
				  					//Se registra el actionRelation (Call)    			         			
				         			Element actionRelation = new Element ("actionRelation");
				         			actionRelation.setAttribute("id", "id."+idCount, xmi);
				        			actionRelation.setAttribute("type", "action:Calls", xmi);
				         			actionRelation.setAttribute("to", "id."+String.valueOf(hsmMethodUnits.get(strNombre)));
	//			    			        			actionRelation.setAttribute("from","id."+String.valueOf(idCount));
				        			actionRelation.setAttribute("from","id."+idAction);
				        			idCount++;			 
				    				
				        			actionElement.addContent(actionRelation);
				        			
				        			//segmento.getChild("model").getChild("codeElement").addContent(codeElement);
				  				
				        		//Si el MethodUnit ya existia	
				  				}else{
				  					
				  					//Se registra el actionRelation (Call)
				         			Element actionRelation = new Element ("actionRelation");
				         			actionRelation.setAttribute("id", "id."+idCount, xmi);
				        			actionRelation.setAttribute("type", "action:Calls", xmi);
				         			actionRelation.setAttribute("to", "id."+String.valueOf(hsmMethodUnits.get(strNombre)));
				        			actionRelation.setAttribute("from","id."+String.valueOf(idAction));
				        			idCount++;
				        			actionElement.addContent(actionRelation);
								
				  				}//FIN Tratamiento la llamada a método
							}//FIN Tratar hijo ASTName
	//TODO Hijos que no sean ASTName (Otro ASTBinOp contatenaciones) 
							else{
System.out.println("ESTO ES UN ASTBinOp distinto.");								
							}
// EJEMPLO ->  AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False" 
						}//FIN Tratamiento tipos de hijos de ASTExpression
					}//FIN Recorrido de hijos de ASTExpression
				}//FIN Tratamiento tipos de hijos de ASTAssignment
			}//FIN Recorrido de hijos de ASTAssignment
		
			if (contextoLocal == ""){
				actionElement.setAttribute("kind", "global");
				segmento.getChild("model").getChild("codeElement").addContent(actionElement);
				this.segmento = segmento;
			}else{
				actionElement.setAttribute("kind", "local");
				Element elementAux = hsmMethodUnitsObj.get(hsmMethodUnits.get(contextoLocal));//.getChild("codeElement");
				elementAux.addContent(actionElement);						
			}		
		
	}
	
	private void trataStatements(ASTStatement astStatement, Element actionElement){
		
		astStatement.getClass().getName();

		//System.out.println("SENTENCIA!!"+astStatement.getClass().getName());    			    			

		/*				
		    			switch (astStatement.getClass().getName()) {

		    			case "ASTIfStatement":
		System.out.println("SENTENCIA IF encontrada!!"+astStatement.getClass().getName());    				
							break;
		    			case "ASTCaseStatment":
		System.out.println("SENTENCIA CASE encontrada!!"+astStatement.getClass().getName());
		    				break;
		    			case "ASTDoWhileStatement":
		System.out.println("SENTENCIA DOWHILE encontrada!!"+astStatement.getClass().getName());
		    				break;
		    			case "ASTForEachStatement":
		System.out.println("SENTENCIA FOREACH encontrada!!"+astStatement.getClass().getName());
		    				break;
		    			case "ASTForWhileWendStatement":
		System.out.println("SENTENCIA FORWHILE encontrada!!"+astStatement.getClass().getName());
		    				break;
		    			case "ASTWithStatement": 
		System.out.println("SENTENCIA WITH encontrada!!"+astStatement.getClass().getName());
		    				break;   			
							
						default:
							break;
						}
		*/
		if(astStatement.getClass().getName().compareTo("cager.parser.ASTIfStatement")==0){

System.out.println("SENTENCIA IF encontrada!!"+astStatement.getClass().getName());
System.out.println(((ASTIfStatement)astStatement).toString());
	    			
			
          
		actionElement.setAttribute("kind", "Condition");
		

		
		if (contextoLocal == ""){
			//actionElement.setAttribute("kind", "global");
			segmento.getChild("model").getChild("codeElement").addContent(actionElement);
			this.segmento = segmento;
		}else{
			//actionElement.setAttribute("kind", "local");
			Element elementAux = hsmMethodUnitsObj.get(hsmMethodUnits.get(contextoLocal));
			elementAux.addContent(actionElement);						
		}
		
		TratarASTIfStatement(astStatement,actionElement);
				
				
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTCaseStatment")==0){
			System.out.println("SENTENCIA CASE encontrada!!"+astStatement.getClass().getName());    		
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTForStatement")==0){
			System.out.println("SENTENCIA FOR encontrada!!"+astStatement.getClass().getName());    		
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTDoWhileStatement")==0){
			System.out.println("SENTENCIA DOWHILE encontrada!!"+astStatement.getClass().getName());   		
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTWhileWendStatement")==0){
			System.out.println("SENTENCIA WHILEWEND encontrada!!"+astStatement.getClass().getName());   		
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTForEachStatement")==0){
			System.out.println("SENTENCIA FOREACH encontrada!!"+astStatement.getClass().getName());    		
		}else if(astStatement.getClass().getName().compareTo("cager.parser.ASTWithStatement")==0){
			System.out.println("SENTENCIA WITH encontrada!!"+astStatement.getClass().getName());   			
		}else{
			System.out.println("SENTENCIA encontrada SIN DETERMINAR!!"+astStatement.getClass().getName());   		
		}
	
	}

	private void TratarASTIfStatement(ASTStatement astStatement, Element actionElement) {

		//Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
//<actionRelation xmi:id="id.26" xmi:type="action:Reads" to="id.20" from="id.25"/>
		Element actionRelation = new Element ("actionRelation");
		actionRelation.setAttribute("id", "id."+idCount, xmi);
		idCount++;
		actionRelation.setAttribute("type", "action:Reads", xmi);
//actionRelation.setAttribute("to", "id."+(idAction+2));//TODO calcular a dónde va
		actionRelation.setAttribute("from","id."+idAction);
			
		actionElement.addContent(actionRelation);
		
		//<actionRelation xmi:id="id.46" xmi:type="action:Writes" to="id.39"/>
		actionRelation = new Element ("actionRelation");
		actionRelation.setAttribute("id", "id."+idCount, xmi);
		idCount++;
		actionRelation.setAttribute("type", "action:TrueFlow", xmi);
//actionRelation.setAttribute("to", "id."+(idAction+2));//TODO calcular a dónde va
		actionRelation.setAttribute("from","id."+idAction);
					 
		actionElement.addContent(actionRelation);
			  				
		actionRelation = new Element ("actionRelation");
		actionRelation.setAttribute("id", "id."+idCount, xmi);
		idCount++;
		actionRelation.setAttribute("type", "action:FalseFlow", xmi);
//actionRelation.setAttribute("to", "id."+(idAction+1));//TODO calcular a dónde va
		actionRelation.setAttribute("from","id."+idAction);
		
		actionElement.addContent(actionRelation);

//		<codeElement xmi:id="id.34" xmi:type="action:ActionElement" name="a3.1" kind="Condition">
//	    	<actionRelation xmi:id="id.35" xmi:type="action:Reads" to="id.29" from="id.34"/>
//	    	<actionRelation xmi:id="id.36" xmi:type="action:TrueFlow" to="id.38" from="id.28"/>
//	    	<actionRelation xmi:id="id.37" xmi:type="action:FalseFlow" to="id.42" from="id.34"/>
//	    </codeElement>

	
//			<codeElement xmi:id="id.111" xmi:type="action:ActionElement" name="p5.2" kind="GreaterThan">
//				<actionRelation xmi:id="id.112" xmi:type="action:Reads" to="id.104" from="id.111"/>
//				<actionRelation xmi:id="id.113" xmi:type="action:TrueFlow" to="id.115" from="id.111"/>
//				<actionRelation xmi:id="id.114" xmi:type="action:FalseFlow" to="id.120" from="id.111"/>
//			</codeElement>
		
		/*
		 
		<codeElement xmi:id="id.25" xmi:type="action:ActionElement" name="1.3" kind="Condition">
			<actionRelation xmi:id="id.26" xmi:type="action:Reads" to="id.20" from="id.25"/>
			<actionRelation xmi:id="id.27" xmi:type="action:TrueFlow" to="id.29" from="id.25"/>
			<actionRelation xmi:id="id.28" xmi:type="action:FalseFlow" to="id.39" from="id.25"/>
		</codeElement>
		
		<codeElement xmi:id="id.75" xmi:type="action:ActionElement" name="a1" kind="Condition">
				<actionRelation xmi:id="id.76" xmi:type="action:TrueFlow" to="id.77" from="id.75" />
				<actionRelation xmi:id="id.77" xmi:type="action:FalseFlow" to="id.76" from="id.75" />
		</codeElement>
		 
		*/
		
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
		
		FileOutputStream fout;
		try {
			fout = new FileOutputStream( ruta+"VBParser\\ejemplosKDM\\archivo.kdm");
			//fout = new FileOutputStream("E:\\WorkspaceParser\\VBParser\\ejemplosKDM\\archivo.kdm");
			xmloutputter.output(documentJDOM, fout);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return docStr;
	}

	private Element inicializaKDM(String segmentName){
    	
    	// Creamos el builder basado en SAX
    	SAXBuilder builder = new SAXBuilder();  
    	
    	//Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
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
		
		segmento.setAttribute("name", segmentName.substring(0, segmentName.indexOf('.')));

		Element modelo = new Element ("model");
		modelo.setAttribute("id", "id.0", xmi);
		idCount++;
		modelo.setAttribute("name", segmentName.substring(0, segmentName.indexOf('.')));
		modelo.setAttribute("type", "code:CodeModel", xmi);
		
		
		Element codeElement = new Element("codeElement");
		codeElement.setAttribute("id", "id.1", xmi);
		idCount++;
		codeElement.setAttribute("name", segmentName);
		codeElement.setAttribute("type", "code:CompilationUnit", xmi);
		
		modelo.addContent(codeElement);
		
		elementoTipos = inicializaDataTypes();
		
		modelo.addContent(elementoTipos);
		
		segmento.addContent(modelo);
		
		
		
		return segmento;
    }
    
	private Element inicializaDataTypes(){
	
		// Creamos el builder basado en SAX
    	SAXBuilder builder = new SAXBuilder(); 
		
    	//Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
    	Namespace code = Namespace.getNamespace("code","http://kdm.omg.org/code");
    	
    	
    	Element model = new Element("model");
    	model.setAttribute("id", "id."+idCount, xmi);
    	model.setAttribute("type", "code:CodeModel", xmi);
		idCount++;
		
		Element codeElementP = new Element("codeElement");
		codeElementP.setAttribute("id", "id."+idCount, xmi);
		codeElementP.setAttribute("type", "code:LanguageUnit", xmi);
		idCount++;
		
		
		Element codeElementI = new Element("codeElement");
		
		for(int i=0;i<dataTypes.length;i++){
			
			codeElementI = new Element("codeElement");
			codeElementI.setAttribute("id", "id."+idCount, xmi);
			if(dataTypes[i].equals("Double")||dataTypes[i].equals("Long")||dataTypes[i].equals("Short")){
				codeElementI.setAttribute("type", "code:DecimalType", xmi);
			}else if(dataTypes[i].equals("Byte")){ 
				codeElementI.setAttribute("type", "code:OctetType", xmi);
			}else{
				codeElementI.setAttribute("type", "code:"+dataTypes[i]+"Type", xmi);
			}
			
			codeElementI.setAttribute("name", dataTypes[i]);
			
			   
			if(!hsmLanguajeUnitDataType.containsKey(dataTypes[i])){
				hsmLanguajeUnitDataType.put(dataTypes[i], idCount);
			}
			
			idCount++;

			codeElementP.addContent(codeElementI);
			
		}
		
			//model.addContent(codeElementP);		
		
    	return codeElementP;
    	
	}

	private void XMLtoJavaParser() {
		
		 // Creamos el builder basado en SAX  
        SAXBuilder builder = new SAXBuilder();  
        
        try {
        	
	        String nombreMetodo = null;
        	List<String> atributos = new ArrayList<String>();
        	String tipoRetorno = null;
        	String nombreArchivo = null;
        	String exportValue = "";
        	String codigoFuente = "";
        	HashMap <String,String> tipos = new HashMap <String,String> ();
        	        	
        	
        	// Construimos el arbol DOM a partir del fichero xml  
	        Document doc = builder.build(new FileInputStream(ruta+"VBParser\\ejemplosKDM\\archivo.kdm"));
			//Document doc = builder.build(new FileInputStream("E:\\WorkspaceParser\\VBParser\\ejemplosKDM\\archivo.kdm"));    
	       
	        Namespace xmi = Namespace.getNamespace("xmi", "http://www.omg.org/XMI");
	        
	        XPathExpression<Element> xpath = XPathFactory.instance().compile("//codeElement", Filters.element());
	    	
	        List<Element> elements = xpath.evaluate(doc);
	    	
	        for (Element emt : elements) {
	        	
	        	if(emt.getAttribute("type",xmi).getValue().compareTo("code:CompilationUnit")==0){
	        	
	        		nombreArchivo = emt.getAttributeValue("name").substring(0, emt.getAttributeValue("name").indexOf('.'));
	        	
	        	}
	        	
	        	if(emt.getAttribute("type",xmi).getValue().compareTo("code:LanguageUnit")==0){
	        	
	        		List<Element> hijos = emt.getChildren();
	        		
	        		for (Element hijo : hijos){
	        			
	        			tipos.put(hijo.getAttributeValue("id",xmi), hijo.getAttributeValue("name"));
	        			
	        		}
	        		
	        	}
	        }
	        
	        
	        FileOutputStream fout;
	        
	        fout = new FileOutputStream(ruta+"VBParser\\src\\cager\\parser\\test\\"+nombreArchivo+".java");	        
	        //fout = new FileOutputStream("E:\\WorkspaceParser\\VBParser\\src\\cager\\parser\\test\\"+nombreArchivo+".java");	        
	        // get the content in bytes
	     	byte[] contentInBytes = null;
	       
	     	contentInBytes=("package cager.parser.test;\n\n").getBytes();
     		
     		fout.write(contentInBytes);
	     	fout.flush();

	     	contentInBytes=("public class "+nombreArchivo+"{\n\n").getBytes();
	     	
	     	fout.write(contentInBytes);
	     	fout.flush();
	     	
	     	 
	        for (Element emt : elements) {
	    	   // System.out.println("XPath has result: " + emt.getName()+" "+emt.getAttribute("type",xmi));
	    	    if(emt.getAttribute("type",xmi).getValue().compareTo("code:MethodUnit")==0){
	    	    	
	    	    	nombreMetodo = emt.getAttribute("name").getValue();
	    	    	
	    	    	if(emt.getAttribute("export")!=null)exportValue = emt.getAttribute("export").getValue();
	    	    	
	    	    	atributos = new ArrayList<String>();
	    	    	
	    	    	List<Element> hijos = emt.getChildren();
	    	    	
	    	    	for (Element hijo : hijos){
	    	    		
	    	    		if (hijo.getAttribute("type",xmi) != null){
	    	    		
		    	    		 if(hijo.getAttribute("type",xmi).getValue().compareTo("code:Signature")==0){
		    	    			 
		    	    			 List<Element> parametros = hijo.getChildren();
		    	    			 
		    	    			 for(Element parametro : parametros){
		    	    				 
		    	    				 if(parametro.getAttribute("kind")==null || parametro.getAttribute("kind").getValue().compareTo("return")!=0){
		    	    					 atributos.add(tipos.get(parametro.getAttribute("type").getValue())+" "+parametro.getAttributeValue("name"));
		    	    				 }else{
		    	    					 tipoRetorno = tipos.get(parametro.getAttribute("type").getValue());
		    	    				 }
		    	    			 
		    	    			 }
		    	    			 
		    	    		 }
	    	    		}else if(hijo.getAttribute("snippet")!=null){
	    	    			
	    	    			codigoFuente = hijo.getAttribute("snippet").getValue(); 
	    	    			
	    	    		}
	    	    		
	    	    	}
	    	    	
	    	    	//System.out.println("MethodUnit!! " + emt.getName()+" "+emt.getAttribute("type",xmi)+" "+emt.getAttribute("name").getValue());
	    	    	//System.out.println(emt.getAttribute("name").getValue());
	    	    	
	    	    	if(tipoRetorno.compareTo("Void")==0){
	    	    		tipoRetorno = "void";
	    	    	}
	    	     			
	    	     	contentInBytes=("\t"+exportValue+" "+tipoRetorno+" "+nombreMetodo+" (").getBytes();
	    	      
	    	     	fout.write(contentInBytes);
	    	     	fout.flush();
	    	     	int n = 0;
	    	     	for(String parametro : atributos){

	    	     		
	    	     		if(atributos.size()>0 && n<atributos.size()-1){
	    	     			contentInBytes=(" "+parametro+",").getBytes();	
	    	     		}else{
	    	     			contentInBytes=(" "+parametro).getBytes();
	    	     		}
	    	     		
	    	     		fout.write(contentInBytes);
		    	     	fout.flush();
		    	     	n++;
	    	     	}
	    	    	
	    	     	contentInBytes=(" ) {\n").getBytes();
    	     		
    	     		fout.write(contentInBytes);
	    	     	fout.flush();
	    	     	
	    	     	contentInBytes=("\n/* \n "+codigoFuente+" \n */\n").getBytes();
    	     		
    	     		fout.write(contentInBytes);
	    	     	fout.flush();

	    	     	contentInBytes=("\t}\n\n").getBytes();
	    	     	
	    	     	fout.write(contentInBytes);
	    	     	fout.flush();
    	     	
	    	    	System.out.print("\t"+exportValue+" "+tipoRetorno+" "+nombreMetodo+" (");
	    	    	n = 0; 
	    	    	for(String parametro : atributos){
	    	    		if(atributos.size()>0 && n<atributos.size()-1){
	    	    			System.out.print(" "+parametro+", ");	
	    	     		}else{
	    	     			System.out.print(" "+parametro);
	    	     		}
	    	    		n++;
	    	    		
	    	    	}
	    	    	System.out.println(" ) {");
	    	    	System.out.println("/* \n "+codigoFuente+" \n */");
	    	    	System.out.println("\t}\n");
	    	    }
	    	    
	    	}
	        
	        contentInBytes=("}\n").getBytes();
     		
     		fout.write(contentInBytes);
	     	fout.flush();
	        fout.close();
	        
	    	XPathExpression<Attribute> xp = XPathFactory.instance().compile("//@*", Filters.attribute(xmi));
	    	for (Attribute a : xp.evaluate(doc)) {
	    	    a.setName(a.getName().toLowerCase());
	    	}
	    	
	        xpath = XPathFactory.instance().compile("//codeElement/@name='testvb.cls'", Filters.element());
	        Element emt = xpath.evaluateFirst(doc);
        	if (emt != null) {
        	    System.out.println("XPath has result: " + emt.getName());
        	}
        
        
        
        } catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JDOMException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
        
	}

}
