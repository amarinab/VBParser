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

public class ParseVBForm extends HttpServlet
{

	public void doPost( HttpServletRequest request, HttpServletResponse response ) throws ServletException, IOException
	{
		response.setContentType("text/html");

		PrintWriter out = response.getWriter();
		HttpSession session = request.getSession(true);

		Integer ival = (Integer)session.getAttribute("counter");
		ival = new Integer( ival == null ? 1: ival.intValue() + 1);

		session.setAttribute("counter", ival);

		String actionString = request.getParameter("action");
		boolean dump = actionString == null ? true : actionString.charAt(0) == 'd';

		out.println("<!DOCTYPE HTML PUBLIC > <HTML> <HEAD>");
		out.println("<TITLE>Example VB Parser output</TITLE>");
		out.println("</HEAD><BODY>");


		if (dump)
        {
			out.println("<H1>Parse Tree</H1>");
            out.println("<PRE>");
            StringReader r = new StringReader(request.getParameter("PROG"));
            VBParser parser = new VBParser(r);
            try
            {
                parser.CompilationUnit().dump(out, "");
            }
            catch (ParseException pe)
            {
                throw new ServletException("Syntax Error",pe); 
            }
        }
		else
        {
			out.println("<H1>Error Handling</H1>");
            out.println("<PRE>");
            out.flush();

            RequestDispatcher d = request.getRequestDispatcher("cager.parser.ParseVB");
            if (d != null)  d.include(request, response);
        }

		out.println("</PRE>");

		out.println("</BODY> </HTML>");
	}



	public String getServletInfo()
	{
		return "VB Parser";
	}

	/**
	 * @uml.property  name="linePos"
	 */
	private int linePos = 0;
	/**
	 * @uml.property  name="tabLength"
	 */
	private int tabLength = 4;

	private void writePlainString(PrintWriter out, String x)
	{
		linePos = 0;

		for( int i = 0 ; i < x.length() ; ++i )
		{
			char ch = x.charAt( i ) ;
			switch( ch )
			{

				case '<' :
					out.print( "&lt;" );
					linePos++;
					break ;

				case '>' :
					out.print( "&gt;" );
					linePos++;
					break ;


				case '\t':
					do
					{
						out.print(" ");
						linePos++;
					}
					while (linePos % tabLength != 0);

					break;

				//case '\r' :
				//	break ;

				case '\n' :
					out.print( ch ) ;
					linePos = 0;
					break ;

				default :
					out.print( ch ) ;
			}
		}
	}

}

