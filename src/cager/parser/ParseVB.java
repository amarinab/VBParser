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

public class ParseVB extends HttpServlet
{

    public void doGet(HttpServletRequest req, HttpServletResponse response) throws ServletException, IOException
    {
        PrintWriter out = response.getWriter();
        out.println("It has done a GET");
        throw new ServletException("A");
    }

    public void doPost( HttpServletRequest request, HttpServletResponse response ) throws ServletException, IOException
    {
        System.out.println("POST ParseVB");

        PrintWriter out = response.getWriter();
        StringReader r = new StringReader(request.getParameter("PROG"));
        BufferedReader in = new BufferedReader(r);

        int options = VBErrHandlerVisitor.SILENT;
        int minLines = 0;
        int handlerType = VBErrHandlerVisitor.SIMPLE_ERR_HANDLER;
        String userHandler = null;
        String selectedProcs = null;
        boolean allProcs = true;
        boolean convertToProps = false;
        String propLet = "P";
        String propGet = "P";

        String value;

        value = request.getParameter("addhandler");
        if (value != null) options |= Integer.parseInt(value);

        value = request.getParameter("linenumbers");
        if (value != null) options |= Integer.parseInt(value);

        value = request.getParameter("minlines");
        if (value != null) minLines = Integer.parseInt(value);
        if (minLines < 1) minLines = 1;

        value = request.getParameter("handlertype");
        if (value != null) handlerType = Integer.parseInt(value);

        if (handlerType == VBErrHandlerVisitor.USER_ERR_HANDLER)
        {
            userHandler = request.getParameter("handlerProg");
        }

        value = request.getParameter("SELECTEDPROCS");
        if (value != null) selectedProcs = value.trim();

        value = request.getParameter("scope");
        if (value != null) allProcs = value.equals("1");

        value = request.getParameter("convertpublics");
        if (value != null) convertToProps = value.equals("1");

        value = request.getParameter("proplet");
        if (value != null) propLet = value;

        value = request.getParameter("propget");
        if (value != null) propGet = value;

        response.setContentType("text/plain");

        try
        {
            VBParser parser = new VBParser(in);
            SimpleNode node = parser.CompilationUnit();

            VBErrHandlerVisitor errVis = new VBErrHandlerVisitor(options, handlerType, minLines, userHandler);

            value = request.getParameter("VBComponent");
            if (value != null)
                errVis.setVBComponent(value);


            value = request.getParameter("VBProject");
            if (value != null)
                errVis.setVBProject(value);

            value = request.getParameter("addtrace");
            if (value != null)
                errVis.setAddTrace(true);

            if (!allProcs && selectedProcs != null && selectedProcs.length() != 0)
                errVis.setSelectedProcedures(selectedProcs);


            node.jjtAccept(errVis, null);

            if (convertToProps)
            {
                ToPropertiesVisitor propVis = new ToPropertiesVisitor();
                propVis.propLet = propLet;
                propVis.propGet = propGet;

                node.jjtAccept(propVis, null);
            }

            out.println("*replace");
            out.print(node.allText(true));

            //System.out.println("POST EXIT WITH");
            //System.out.println(node.allText(true));
        }
        catch(Exception e)
        {
            response.reset();

            out.println("*error");
            out.println(e);
            e.printStackTrace(out);
            out.println("Prog");
            out.println(request.getParameter("PROG"));
        }
        catch (Throwable t)
        {
            response.reset();

            out.println("*error");
            out.println(t);
        }
    }



    public String getServletInfo()
    {
        return "VB Parser";
    }
}

