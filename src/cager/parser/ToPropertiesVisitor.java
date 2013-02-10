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
package cager.parser;

import java.io.*;
import java.util.*;
import cager.parser.*;



public class ToPropertiesVisitor extends VBParserVisit implements VBParserTreeConstants, VBParserConstants
{
	public void start()
	{
	}

	/**
	 * @uml.property  name="propLet"
	 */
	public String propLet = "P";
	/**
	 * @uml.property  name="propGet"
	 */
	public String propGet = "P";

	/**
	 * @uml.property  name="inDecl"
	 */
	private boolean inDecl;
	/**
	 * @uml.property  name="name"
	 * @uml.associationEnd  
	 */
	private ASTName name;
	/**
	 * @uml.property  name="typeName"
	 * @uml.associationEnd  
	 */
	private ASTTypeName typeName;

	private ASTStatement compileStatement(String src) throws ParseException, IOException
	{
		VBParser p = new VBParser(new StringReader(src));
		return p.Statement();
	}

	private ASTProcDeclaration compileProcDelcaration(String src) throws ParseException, IOException
	{
		// Compile a CompilationUnit rather than a ProcDeclaration, so that we may include
		// the trailing new lines etc.

		VBParser p = new VBParser(new StringReader(src));
		ASTCompilationUnit cu = p.CompilationUnit();
		if (cu.children.length != 1)
			throw new ParseException("Invalid Procedure: " + src);

		// Now adjust the start and end tokens of the ProcDeclaration to be the same as
		// the Compilation unit (horrible bodge!!)
		((SimpleNode)cu.children[0]).begin = cu.begin;
		((SimpleNode)cu.children[0]).end = cu.end;

		return (ASTProcDeclaration)(cu.children[0]);
	}

	public Object visit(ASTCompilationUnit node, Object data)
	{
		if (node.children == null)
			return data;

		if (propLet.equals("P"))
			propLet = "Public";
		else if (propLet.equals("F"))
			propLet = "Friend";
		else if (propLet.equals("N"))
			propLet = null;

		if (propGet.equals("P"))
			propGet = "Public";
		else if (propGet.equals("F"))
			propGet = "Friend";

		for (int i = 0; i < node.children.length; i++)
		{
			if (node.children[i] instanceof ASTDeclaration)
			{
				ASTDeclaration decl = (ASTDeclaration)node.children[i];

				inDecl = true;
				name = null;
				typeName = null;
				decl.childrenAccept(this, data);
				inDecl = false;

				if (decl.begin.image.equalsIgnoreCase("public") || decl.begin.image.equalsIgnoreCase("friend"))
				{
					//System.out.println("Got a public " + name.allText() + "/" + name.getName());
					decl.begin.image = "Private";

					String oldName = name.getName();
					String newName = "m_" + name.getName();
					String typeNameStr = (typeName == null ? "Variant" : typeName.allText());
					// Guess if the TypeName refers to an object. It is NOT an
					// object if it is a VB base type (String, Integer etc), otherwise we guess
					// that all class names begin with an upper case, and all enums begin with a
					// lower case.
					boolean isObject = !typeName.isBaseType() && Character.isUpperCase(typeName.allText().charAt(0));

					name.setName(newName);

					try
					{
						if (propGet != null)
						{
							String getStr =
								"\n" + propGet + " Property Get " + oldName + "() As " + typeNameStr + "\n" +
								"    " + (isObject ? "Set " : "") + oldName + " = " + newName + "\n" +
								"End Property\n";


							ASTProcDeclaration get = compileProcDelcaration(getStr);

							node.insertChild(get, -1);
						}

						if (propLet != null)
						{
							String letStr =
								"\n" + propLet + " Property " + (isObject ? "Set " : "Let") + " " + oldName + "(Byval RHS As " + typeNameStr + ")" + "\n" +
								"    " + (isObject ? "Set " : "") + newName + " = RHS\n" +
								"End Property\n";

							ASTProcDeclaration let = compileProcDelcaration(letStr);


							node.insertChild(let, -1);
						}
					}
					catch (Exception e)
					{
						// Should never happen -- we control the source

						throw new Error("pe " + e.toString());
					}
				}
			}
		}

		return data;
	}


	public Object visit(ASTName node, Object data)
	{
		name = node;
		return data;
	}

	public Object visit(ASTTypeName node, Object data)
	{
		typeName = node;
		return data;
	}



	public static void main(String[] args) throws Exception
	{
		ToPropertiesVisitor v = new ToPropertiesVisitor();

		for (int i = 0; i < args.length; i++)
		{
			v.process(new File(args[i]));
		}
	}

	private void process(File f) throws IOException, ParseException
	{
		System.out.println("Process " + f.getAbsolutePath());

		if (f.isDirectory())
		{
			File[] files = f.listFiles(new VBFileFilter());
			if (files != null)
			{
				for (int i = 0; i < files.length; i++)
					process(files[i]);
			}
		}
		else
		{
			parse(f);
		}
	}

	private void parse(File in) throws IOException, ParseException
	{
		VBParser parser = new VBParser(new FileInputStream(in));
		ASTCompilationUnit node = parser.CompilationUnit();
		//node.dump(">");
		//Token t = node.getFirstToken();
		//while (t != null)
		//{
		//	System.out.println(t.debug());
		//	t = t.next;
		//}

		node.jjtAccept(this, null);

		//node.dump("!");

		System.out.println(node.allText(true));
	}

	private class VBFileFilter implements FileFilter
	{
		public boolean accept(File pathname)
		{
			if (pathname.isDirectory())
				return true;

			String n = pathname.getName().toLowerCase();

			return n.endsWith(".cls") || n.endsWith(".bas") || n.endsWith(".frm");
		}
	}
}

