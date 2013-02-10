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

import java.io.InputStream;

import junit.framework.Assert;
import junit.framework.TestCase;
import cager.parser.SimpleNode;
import cager.parser.VBErrHandlerVisitor;
import cager.parser.VBParser;
import cager.parser.VBParserVisitor;

/**
 * @author paul
 *
 */
public class TestParser extends TestCase
{
    public TestParser(String name)
    {
        super(name);
    }

    public void setUp()
    {
    }

    private void simpleTest() throws Exception
    {
        VBParser parser;
        InputStream is = getClass().getResourceAsStream("testvb.cls");
        Assert.assertNotNull("InputStream null", is);
        
        parser = new VBParser(is);
        parser.setAttemptErrorRecovery(true);
        SimpleNode node = parser.CompilationUnit(false);
        //node = (SimpleNode)(parser.ExprTest());
        //          PrintWriter ostr = new PrintWriter(new FileWriter(args[1]));

        //        node.jjtAccept(new VBErrHandlerVisitor(), null);

        Assert.assertNotNull("node", node);
        
        VBParserVisitor v = new VBErrHandlerVisitor();
        node.jjtAccept(v, null);
        
//        node.dump(">");
//
        System.out.println(node.allText(true));
        System.out.println("VBParser:  Parsing completed successfully.");

        System.out.println("Completed test of generated Parser");
    }
    
    public void testVBParser() throws Exception
    {
        simpleTest();
    }

    public static void main(String[] args) throws Exception
    {
	    System.exit(0);
        new TestParser("main test").simpleTest();
	    System.out.println("Finished");
    }
}
