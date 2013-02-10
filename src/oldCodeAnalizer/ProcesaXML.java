
package oldCodeAnalizer;  
  
import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
import java.util.List;  
  
import org.jdom2.Document;  
import org.jdom2.Element;  
import org.jdom2.JDOMException;  
import org.jdom2.Namespace;  
import org.jdom2.input.SAXBuilder;  
import org.jdom2.output.Format;  
import org.jdom2.output.XMLOutputter;  
import org.jdom2.xpath.XPath;  
  
/** 
 * Clase de ejemplo de procesado de XML mediante JDOM. 
 *  
 * @author Xela 
 *  
 */  
public class ProcesaXML {  
  
   public static void main(String[] args) {  
  
      try {     
         // Creamos el builder basado en SAX  
         SAXBuilder builder = new SAXBuilder();  
         // Construimos el arbol DOM a partir del fichero xml  
         Document doc = builder.build(new FileInputStream("/ruta_a_fichero/fichero.xml"));  
  
         // Obtenemos la etiqueta raíz  
         Element raiz = doc.getRootElement();  
         // Recorremos los hijos de la etiqueta raíz  
         List<Element> hijosRaiz = raiz.getChildren();  
         for(Element hijo: hijosRaiz){  
            // Obtenemos el nombre y su contenido de tipo texto  
            String nombre = hijo.getName();  
            String texto = hijo.getValue();  
              
            System.out.println("\nEtiqueta: "+nombre+". Texto: "+texto);  
              
            // Obtenemos el atributo id si lo hubiera  
            String id = hijo.getAttributeValue("id");  
            if(id!=null){  
               System.out.println("\tId: "+id);  
            }  
         }  
           
         // Obtenemos una etiqueta hija del raiz por nombre  
         Element etiquetaHija = raiz.getChild("etiquetaHija");  
         System.out.println(etiquetaHija.getName());  
         // Incluso podemos obtener directamente el texto de una etiqueta hija  
         String texto = raiz.getChildText("etiquetaHija");  
         System.out.println(texto);  
                    
         // Obtenemos una etiqueta hija del raiz por nombre con Namespaces  
         // Primero creamos el objeto Namespace  
         Namespace nsXela = Namespace.getNamespace("xela", "http://www.latascadexela.es");  
         // Ahora obtenemos el hijo  
         Element etiquetaNamespace = raiz.getChild("etiquetaConNamespace", nsXela);  
         System.out.println(etiquetaNamespace.getName());  
           
         // Buscamos una etiqueta mediante XPath  
         Element etiquetaHijaXP = (Element)XPath.selectSingleNode(doc, "/etiquetaPrincipal/etiquetaHija");  
         System.out.println(etiquetaHijaXP.getName());  
           
         // Buscamos una etiqueta con namespace mediante XPath  
         Element etiquetaNamespaceXP = (Element)XPath.selectSingleNode(doc, "/etiquetaPrincipal/xela:etiquetaConNamespace");  
         System.out.println(etiquetaNamespaceXP.getName());  
           
         // Si hacemos uso muchas veces del mismo XPath sobre varios document  
         // es más eficiente crear un objeto XPath y usarlo varias veces  
         XPath xpathEtiquetaHija= XPath.newInstance("/etiquetaPrincipal/etiquetaHija");  
         Element etiqueta = (Element)xpathEtiquetaHija.selectSingleNode(doc);  
         System.out.println(etiqueta.getName());  
  
         // Creamos una nueva etiqueta  
         Element etiquetaNueva = new Element("etiquetaNueva");  
         // Añadimos un atributo  
         etiquetaNueva.setAttribute("atributoNuevo", "Es un nuevo atributo");  
         // Añadimos contenido  
         etiquetaNueva.setText("Contenido dentro de la nueva etiqueta");  
         // La añadimos como hija a una etiqueta ya existente  
         etiquetaHija.addContent(etiquetaNueva);  
           
         // Vamos a crear un XML desde cero  
         // Lo primero es crear el Document  
         Document docNuevo = new Document();  
         // Vamos a generar la etiqueta raiz  
         Element eRaiz = new Element("raiz");  
         // y la asociamos al document  
         docNuevo.addContent(eRaiz);  
           
         // Vamos a copiar la etiquetaHija del primer document a este  
         // Lo primero es crear una copia de etiquetaHija  
         Element copiaEtiquetaHija = (Element)etiquetaHija.clone();  
         // Después la colocamos como hija de la etiqueta raiz  
         eRaiz.addContent(copiaEtiquetaHija);  
           
         // Vamos a mover la etiquetaConNamespace a este document  
         // Primero la desasociamos de su actual padre  
         etiquetaNamespace.detach();  
         // Una vez que ya es huerfana la podemos colocar donde queramos  
         // Por ejemplo, bajo la etiqueta raiz  
         eRaiz.addContent(etiquetaNamespace);  
           
           
         // Vamos a serializar el XML  
         // Lo primero es obtener el formato de salida  
         // Partimos del "Formato bonito", aunque también existe el plano y el compacto  
         Format format = Format.getPrettyFormat();  
         // Creamos el serializador con el formato deseado  
         XMLOutputter xmloutputter = new XMLOutputter(format);  
         // Serializamos el document parseado  
         String docStr = xmloutputter.outputString(doc);  
         // Serializamos nuestro nuevo document  
         String docNuevoStr = xmloutputter.outputString(docNuevo);  
           
         System.out.println("XML parseado:\n"+docStr);  
         System.out.println("XML nuevo:\n"+docNuevoStr);  
           
      } catch (FileNotFoundException e) {  
         e.printStackTrace();  
      } catch (JDOMException e) {  
         e.printStackTrace();  
      } catch (IOException e) {  
         e.printStackTrace();  
      }  
  
   }  
  
}  
