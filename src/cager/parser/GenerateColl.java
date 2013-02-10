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
import javax.servlet.*;
import javax.servlet.http.*;

public class GenerateColl extends HttpServlet
{

	public void doPost( HttpServletRequest request, HttpServletResponse response ) throws ServletException, IOException
	{
		System.out.println("POST GenerateColl");

		PrintWriter out = response.getWriter();
		String value;
		String collName = "Ccoll";
		String itemName = "Citem";
		boolean dalExtensions = false;

		value = request.getParameter("COLLNAME");
		if (value != null) collName = value;

		value = request.getParameter("ITEMNAME");
		if (value != null) itemName = value;

		value = request.getParameter("DALEXTENSIONS");
		if (value != null) dalExtensions = value.equals("Y");

		response.setContentType("text/plain");


		out.println("*addfile " + collName + ".cls");
		out.println("VERSION 1.0 CLASS");
		out.println("BEGIN");
		out.println("  MultiUse = -1  'True");
		out.println("  Persistable = 0  'NotPersistable");
		out.println("  DataBindingBehavior = 0  'vbNone");
		out.println("  DataSourceBehavior  = 0  'vbNone");
		out.println("  MTSTransactionMode  = 0  'NotAnMTSObject");
		out.println("END");
		out.println("Attribute VB_Name = \"" + collName + "\"");
		out.println("Attribute VB_GlobalNameSpace = False");
		out.println("Attribute VB_Creatable = True");
		out.println("Attribute VB_PredeclaredId = False");
		out.println("Attribute VB_Exposed = True");
		out.println("Option Explicit");
		out.println("");
		out.println("Private col" + " As Collection           ' of C" + itemName + "");
		out.println("");

		if (dalExtensions)
		{
			out.println("Private kindex As Integer               ' stores the next key number");
			out.println("");
			out.println("Public Sub Add(Item As " + itemName + ", Optional AddAsUnchanged As Boolean = False)");
			out.println("    If AddAsUnchanged Then");
			out.println("        Item.State = eDataState.Unchanged");
			out.println("    Else");
			out.println("        Item.State = eDataState.Newdata");
			out.println("    End If");
			out.println("    Item.Key = \"K\" & kindex");
			out.println("    kindex = kindex + 1");
			out.println("    col" + ".Add Item, Item.Key");
			out.println("End Sub");
			out.println("");
			out.println("Public Sub Remove(ByVal Index As Variant)");
			out.println("    If col" + ".Item(Index).State = Newdata Then");
			out.println("        col" + ".Remove Index");
			out.println("    Else");
			out.println("        col" + ".Item(Index).State = Deleted");
			out.println("    End If");
			out.println("End Sub");
			out.println("");
		}
		else
		{
			out.println("Public Sub Add(Item As " + itemName + ")");
			out.println("    col" + ".Add Item");
			out.println("End Sub");
			out.println("");
			out.println("Public Sub Remove(ByVal Index As Variant)");
			out.println("    col" + ".Remove Index");
			out.println("End Sub");
			out.println("");
		}

		out.println("Public Function Item(ByVal Index As Variant) As " + itemName + "");
		out.println("Attribute Item.VB_UserMemId = 0");
		out.println("    Set Item = col" + ".Item(Index)");
		out.println("End Function");
		out.println("");
		out.println("'This function is provided to support the For Each....Next construct when iterating the collection.");
		out.println("Public Function NewEnum() As IUnknown");
		out.println("Attribute NewEnum.VB_UserMemId = -4");
		out.println("    Set NewEnum = col" + ".[_NewEnum]");
		out.println("End Function");
		out.println("");
		out.println("Public Function Count() As Long");
		out.println("    Count = col" + ".Count");
		out.println("End Function");
		out.println("");

		if (dalExtensions)
		{
			out.println("Public Sub Destroy()");
			out.println("    ' method to destroy the contents of the collection; it is not enough simply");
			out.println("    ' to set the collection reference to nothing since there may be another");
			out.println("    ' reference to it; set collection=nothing only reduces the reference count");
			out.println("    If Not col" + " Is Nothing Then");
			out.println("        Do While col" + ".Count");
			out.println("            col" + ".Remove 1");
			out.println("        Loop");
			out.println("    End If");
			out.println("End Sub");
			out.println("");
		}

		out.println("'********************************************************************");
		out.println("' Class initialize and terminate routines                           *");
		out.println("'********************************************************************");
		out.println("");
		out.println("'Initialises the object by instantiating the underlying VB collection which provides much of the functionality of this class");
		out.println("Private Sub Class_Initialize()");
		out.println("    Set col" + " = New Collection");
		out.println("End Sub");
		out.println("");
		if (dalExtensions)
		{
			out.println("'********************************************************************");
			out.println("' Dump routine (for debug)                                          *");
			out.println("'********************************************************************");
			out.println("");
			out.println("'This function returns a string which describes all the properties of the data collection class.  Included are the Protected and Private properties which can be used to assist with debugging.  This routine is primarily intended for use during development of client applications.");
			out.println("Public Function Dump(Optional indent As Integer = 0) As String");
			out.println("    Dim tabs As String");
			out.println("    Dim Item As " + itemName + "");
			out.println("    tabs = Space(indent * 2)");
			out.println("    Dump = _");
			out.println("        tabs & vbCrLf & \"" + collName + " Collection Dump\" & vbCrLf & _");
			out.println("        tabs & \"======================\" & vbCrLf");
			out.println("        ");
			out.println("    For Each Item In col" + "");
			out.println("        Dump = Dump & Item.Dump(indent + 1) & vbCrLf");
			out.println("    Next");
			out.println("End Function");
		}
	}



	public String getServletInfo()
	{
		return "Coll Generator";
	}
}

