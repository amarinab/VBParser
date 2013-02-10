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
import cager.parser.*;

public class TestParser
{
	/**
	 * @uml.property  name="parser"
	 * @uml.associationEnd  
	 */
	private VBParser parser = null;

	private class VBFileFilter implements FileFilter
	{
		public boolean accept(File pathname)
		{
			if (pathname.isDirectory())
				return false;

			String n = pathname.getName().toLowerCase();

			return n.endsWith(".cls") || n.endsWith(".bas") || n.endsWith(".frm");
		}
	}

	private class DirectoryFilter implements FileFilter
	{
		public boolean accept(File pathname)
		{
			return pathname.isDirectory();
		}
	}

	public void runTest(String prefix, File directory)
	{
		ASTCompilationUnit node;

		File[] files = directory.listFiles(new VBFileFilter());

        if (files != null)
        {
			for (int i = 0; i < files.length; i++)
			{
				String fileName = files[i].getAbsolutePath();

				System.out.println(fileName);

				try
				{
					if (parser == null)
					{
						parser = new VBParser(new FileInputStream(files[i]));
					}
					else
					{
						parser.ReInit(new FileInputStream(files[i]));
					}

					node = parser.CompilationUnit(true);
					//node.dump(">");
				}
				catch (FileNotFoundException e)
				{
					System.out.println(files[i].getAbsolutePath() + ": "  + e);
				}
				catch (ParseException e)
				{
					System.out.println(files[i].getAbsolutePath() + ": "  + e);
				}
			}
		}

		File[] subDirs = directory.listFiles(new DirectoryFilter());

		if (subDirs != null)
		{
			for (int i = 0; i < subDirs.length; i++)
			{
				String dirName = subDirs[i].getName();
				//System.out.println(prefix + "Directory " + dirName);
				runTest(prefix + "  ", subDirs[i]);
			}
	}
	}



	public static void main(String args[]) throws Exception
	{
		TestParser t = new TestParser();

		if (args.length == 0)
		{
			File f = new File("C:/Desktop");
			if (f.exists())
			{
				t.runTest("", f);		// At work
			}
			else
			{
				t.runTest("", new File("e:/VBSrc"));	// At home
			}

		}
		else
		{
			for (int i = 0; i < args.length; i++)
				t.runTest("", new File(args[i]));
		}
	}
}

