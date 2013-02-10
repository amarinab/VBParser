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
import java.text.MessageFormat;


/**
**  <P>
**  VBErrHandlerVisitor processes a VB source file to add error handling
**  statements.
**  </P>
*/

public class VBErrHandlerVisitor extends VBParserVisit implements VBParserTreeConstants, VBParserConstants
{
    /**
    **  Produces a report (to <I>stdout</I>) without
    **  creating an output file.
    */
    public static final int ADD_HANDLER             = 0x0001;

    /**
    **  If this flag is set, then any existing error handler is deleted
    **  and replaced by a new one.
    **  </I>Not currently supported.</I>
    */
    public static final int REPLACE_EXISTING        = 0x0002;

    /**
    **  Does not write progress messages to <I>stdout</I>.
    */
    public static final int SILENT                  = 0x0004;

    /**
    **  Apply Line Numbers to VB statements.
    */
    public static final int APPLY_LINE_NUMBERS      = 0x0008;

    /**
    **  Remove Line Numbers from VB statements.
    */
    public static final int REMOVE_LINE_NUMBERS     = 0x0010;

    /**
    **  Add a Trace Line at the start of each procedure.
    */
    public static final int TRACE_ENTRY     = 0x0020;

    /**
	 * Options: ReportOnly, ReplaceExisting, Silent, Apply / Remove line numbers, Add Trace
	 * @uml.property  name="options"
	 */
    private int options = 0;




    /**
    *   A simple Err.Raise Error Handler
    */
    public static final int SIMPLE_ERR_HANDLER      = 1;

    /**
    *   Error Handler that calls a function
    */
    public static final int CALL_ERR_HANDLER        = 2;

    /**
    *   User-Supplied Error Handler.
    */
    public static final int USER_ERR_HANDLER        = 3;

    /**
	 * Type of Error Handler to be generated
	 * @uml.property  name="errorHandlerType"
	 */

    private int errorHandlerType = SIMPLE_ERR_HANDLER;



    /**
	 * @uml.property  name="errorHandlerFormat"
	 */
    private String errorHandlerFormat = "Err.Raise Err.Number, Err.Source, Err.Description & \" (\" & {0} & \")\"";

    /**
	 * A count of the number of statements in a procedure. This count is displayed in the report generated We only insert error handlers into procedures with more than a certain number of lines.
	 * @uml.property  name="statementCount"
	 */
    private int statementCount;

    /**
	 * Only insert error handlers if there are
	 * @uml.property  name="minStatements"
	 */
    private int minStatements = 1;

    /**
	 * @uml.property  name="selectedProcedures"
	 * @uml.associationEnd  multiplicity="(0 -1)" elementType="java.lang.String"
	 */
    private Collection selectedProcedures = null;

    /**
	 * @uml.property  name="vbProject"
	 */
    private String vbProject = "";
    /**
	 * @uml.property  name="vbComponent"
	 */
    private String vbComponent = "";

    /**
    **  Equivalent to <code>VBErrHandlerVisitor(ADD_HANDLER, SIMPLE_ERR_HANDLER, 0)</code>
    */
    public VBErrHandlerVisitor()
    {
        this(ADD_HANDLER, SIMPLE_ERR_HANDLER);
    }

    /**
    **  Equivalent to <code>VBErrHandlerVisitor(options, errorHandlerType, 0)</code>
    */
    public VBErrHandlerVisitor(int options, int errorHandlerType)
    {
        this(options, errorHandlerType, 1);
    }

    /**
    **  Equivalent to <code>VBErrHandlerVisitor(options, USER_ERR_HANDLER, 0, errorHandlerFormat)</code>
    */
    public VBErrHandlerVisitor(int options, String errorHandlerFormat)
    {
        this(options, USER_ERR_HANDLER, 1, errorHandlerFormat);
    }

    public VBErrHandlerVisitor(int options, int errorHandlerType, int minStatements)
    {
        this(options, errorHandlerType, minStatements, null);
    }

    /**
    **  Constructs a VBErrHandlerVisitor with the specified options, and which
    **  will write the modified source to the specified file.
    **
    **  @param   options        Silent, ReportOnly, ReplaceExisting.
    **  @param   errorHandlerType   SIMPLE_ERR_HANDLER or CALL_ERR_HANDLER
    */
    public VBErrHandlerVisitor(int options, int errorHandlerType, int minStatements, String errorHandlerFormat)
    {
        super();
        this.options = options;
        this.errorHandlerType = errorHandlerType;
        this.minStatements = minStatements;
        this.errorHandlerFormat = errorHandlerFormat;
    }


    public void setSelectedProcedures(String list)
    {
        if (list == null)
        {
            selectedProcedures = null;
        }
        else
        {
            StringTokenizer tok = new StringTokenizer(list, ",");
            java.util.Collection coll = new java.util.Vector();

            while (tok.hasMoreTokens())
            {
                coll.add(tok.nextToken());
            }

            selectedProcedures = coll;
        }
    }

    public void setVBComponent(String vbComponent)
    {
        this.vbComponent = vbComponent;
    }

    public void setVBProject(String vbProject)
    {
        this.vbProject = vbProject;
    }

    public void setAddTrace(boolean value)
    {
        if (value)
            this.options |= TRACE_ENTRY;
        else
            this.options &= ~TRACE_ENTRY;
    }


    public void start()
    {
        startLineNumber = 10;
    }

    public Object visit(ASTProcDeclaration node, Object data)
    {
        System.out.println("Visit ProcDecl");

        if (selectedProcedures == null || selectedProcedures.contains(node.getName()))
        {
            try
            {
                String procName = node.getName();

                statementCount = 0;
                node.childrenAccept(this, data);

                ASTStatements statements = node.getStatements();

                ASTOnError onErr = (ASTOnError)(node.findNode(JJTONERROR));

                System.out.println("!!!!!Proc \"" + procName + "\": " + statementCount + " " + (onErr == null ? "No Handler" : onErr.getGotoLabel().getName()));

                if ((options & SILENT) == 0)
                {
                    System.out.println("Proc \"" + procName + "\": " + statementCount + " " + (onErr == null ? "No Handler" : onErr.getGotoLabel().getName()));
                }

                if ((options & TRACE_ENTRY) != 0)
                {
                    if (statementCount >= minStatements)
                        addTrace(node);
                }

                // The label will normally be a alphanumeric label, but could be a
                // line number. Therefore use SimpleNode, as it might return either
                // a ASTLabel or a ASTLiteral.
                //SimpleNode errHandler = null;

                if (onErr == null)
                {
                    // No "On Error"

                    if ((options & ADD_HANDLER) != 0)
                    {
                        // Do not create error handlers for single statement (or empty procedures)..
                        // This should be configurable.
                        if (statementCount >= minStatements)
                            addErrorHandler(node);
                    }
                }
                else if (onErr.getGotoLabel() == null)
                {
                    // Must have been "On Error Resume Next" etc.
                    /*
                    ** TODO - Might be better to generate a wrapper error
                    ** handler here.
                    */
                }
                else if (! (onErr.getGotoLabel() instanceof ASTName))
                {
                    /*
                    **  Maybe we should handle this, but surely no-one still uses
                    **  line numbers in this way.
                    */

                    System.out.println("On Error Goto Line Number not supported");
                }
                else
                {
                    // On Error Goto Label

                    ASTName label = (ASTName)onErr.getGotoLabel();
                }

                if ((options & APPLY_LINE_NUMBERS) != 0)
                {
                    applyLineNumbers(statements, false);
                }
                else if ((options & REMOVE_LINE_NUMBERS) != 0)
                {
                    applyLineNumbers(statements, true);
                }
            }
            catch(Exception e)
            {
                System.out.println("Error adding ErrorHandler");
                System.out.println(e);
                e.printStackTrace();
            }
        }

        return data;
    }

    // This version of "visit" is called for ASTStatement and all of its derived
    // classes (this is done by the VBParserVisit class).
    // We just count the number of statements in each class.
    public Object visit(ASTStatement node, Object data)
    {
        //System.out.println("Statement: " + node);
        statementCount++;
        node.childrenAccept(this, data);
        return data;
    }


    private void addErrorHandler(ASTProcDeclaration proc) throws ParseException
    {
        ASTStatements statements = proc.getStatements();

        String procType;
        switch (proc.getProcType())
        {
            case SUB: procType = "Sub"; break;
            case FUNCTION: procType = "Function"; break;
            case PROPERTY: procType = "Property"; break;
            default: procType = "?";
        }


        System.out.println("Adding Error Handling");
        //System.out.println("BEFORE");
        //System.out.println(statements.allText(true));

        statements.insertChild(new ASTStatement(JJTSTATEMENT), 0, "On Error Goto Handler\r\n");

        if (errorHandlerType == SIMPLE_ERR_HANDLER)
        {
            statements.insertChild(
                    makeStatements(new String[] {
                        "Exit " + procType,
                        "Handler:",
                        "      Err.Raise Err.Number, Err.Source, \"" + proc.getName() + ": \" & Err.Description"}),
                    -1);
        }
        else if (errorHandlerType == USER_ERR_HANDLER)
        {
            Object[] args = {
                proc.getName(),
                vbComponent,
                vbProject
                };

            statements.insertChild(
                    makeStatements(new String[] {
                        "Exit " + procType,
                        "Handler:",
                        MessageFormat.format(errorHandlerFormat, args)}),
                    -1);

        }
        else
        {
            statements.insertChild(
                    makeStatements(new String[] {
                        "Exit " + procType,
                        "Handler:",
//                      "      Dim ErrNumber as Long: ErrNumber = Err.Number",
//                      "      Dim ErrSource as String: ErrSource = Err.Source",
//                      "      Dim ErrDescription as String: ErrDescription = Err.Description",
                        "      HandleError Err.Number, Err.Source, Err.Description, \"" + proc.getName()  + "\""
                        }),
                    -1);
        }

        //System.out.println("AFTER");
        //System.out.println(statements.allText(true));
    }


    private void addTrace(ASTProcDeclaration proc) throws ParseException
    {
        ASTStatements statements = proc.getStatements();

        statements.insertChild(new ASTStatement(JJTSTATEMENT), 0, "Trace(\"Entered " + proc.getName() + "\")\r\n");
    }


    private SimpleNode[] findErrDesc(SimpleNode n)
    {
        SimpleNode[] nodes = n.findNodes(JJTBINOP, ".");
        Vector v = new Vector(nodes.length);

        System.out.println("Found " + nodes.length + " binary dots");

        for (int i = 0; i < nodes.length; i++)
        {
            SimpleNode lhs = (SimpleNode)nodes[i].jjtGetChild(0);
            SimpleNode rhs = (SimpleNode)nodes[i].jjtGetChild(1);

            if (lhs.id == JJTNAME && lhs.getFirstToken().image.equalsIgnoreCase("Err") && rhs.id == JJTNAME && rhs.getFirstToken().image.equalsIgnoreCase("Description"))
            {
                ((SimpleNode)nodes[i]).dump("ED ");
                v.add(nodes[i]);
            }
        }

        return (SimpleNode[])v.toArray(new SimpleNode[v.size()]);
    }

    private boolean expressionHasErl(SimpleNode n)
    {
        while (! (n instanceof ASTExpression) )
            n = (SimpleNode)n.jjtGetParent();

        SimpleNode[] erlNodes = n.findNodes(JJTNAME, "Erl");

        return erlNodes.length > 0;
    }

    private ASTExpression compileExpression(String src) throws ParseException, IOException
    {
        VBParser p = new VBParser(new StringReader(src));
        return p.Expression();
    }

    private void addErl(SimpleNode errDesc) throws ParseException, IOException
    {
        // errDesc is the BinaryOp for "Err.Description"
        SimpleNode parent = (SimpleNode)errDesc.jjtGetParent();

        ASTExpression expr = compileExpression(" Err.Description & \"(\" & Erl & \")\" ");

        expr.dump("New Expr");

        ASTBinOp exprPart = (ASTBinOp)(expr.jjtGetChild(0));

        parent.replaceChild(errDesc, exprPart);
        parent.dump("New Parent");
    }

    private ASTStatement[] makeStatements(String[] text)
    {
        ASTStatement[] statements = new ASTStatement[text.length];

        for (int i = 0; i < text.length; i++)
        {
            Token t = Token.newToken(IDENTIFIER);
            t.image = text[i] + "\r\n";
            t.next = null;
            t.specialToken = null;
            ASTStatement s = new ASTStatement(JJTSTATEMENT);
            s.setFirstToken(t);
            s.setLastToken(t);

            statements[i] = s;

            if (i > 0)
                statements[i - 1].getLastToken().next = t;
        }

        return statements;
    }


    /**
	 * @uml.property  name="startLineNumber"
	 */
    private int startLineNumber;

    /**
    **  Apply / remove line numbers to a block of statements.
    **  This function loops around all of the tokens in a statement, adding a
    **  LINE_NUMBER special token to the start of each line. (An alternative
    **  implemenation would be to loop around each statement in the block, adding
    **  the LINE_NUMBER to the statement. This was rejected, as it would not cope with
    **  more than one statement per line).
    */
    // TODO -- how to return line number; make this proc take a ASTProcDeclaration[].
    // might be better to make startLineNumber class-level.
    private void applyLineNumbers(ASTStatements statements, boolean remove) throws ParseException, IOException
    {
        statements.dump("BEFORE");

        SimpleNode[] errDescs = findErrDesc(statements);

        for (int i = 0; i < errDescs.length; i++)
        {
            if (!expressionHasErl(errDescs[i]))
            {
                addErl(errDescs[i]);
            }
        }

        statements.dump("AFTER");

        Token t = statements.getFirstToken();

        if (t == null)
        {
            // An empty statement block
            return;
        }

        boolean atBOL = true;

        // TODO - tidy up EOL handling
        while (t != statements.getLastToken())
        {
            if (atBOL)
            {
                if (t.kind != EOL)
                {
                    // First token of a line. Does it already have a line number? Search
                    // backwards for a LINE_NUMBER. A BLANK_LINE token means that we were preceded
                    // by a blank line -- stop at one of these as well.

                    atBOL = false;

                    Token special = t.specialToken;

                    while (special != null && special.kind != LINE_NUMBER && special.kind != BLANK_LINE)
                        special = special.specialToken;


                    if (remove)
                    {
                        if (special == null || special.kind == BLANK_LINE)
                        {
                            // No line number - nothing to do
                        }
                        else
                        {
                            // Remove it and the 3 following spaces.
                            //System.out.println("REMOVE LINE NU: " + special.debug());
                            //System.out.println("              : " + special.next.debug());

                            special.image = "";
                            if (special.next != null & special.next.image.substring(0, 3).equals("   "))
                                special.next.image = special.next.image.substring(3);
                        }
                    }
                    else
                    {
                        if (special == null || special.kind == BLANK_LINE)
                        {
                            // No line number - add it
                            Token newToken = new Token(LINE_NUMBER);
                            newToken.image = Integer.toString(startLineNumber) + "   ";

                            // Now link to the beginning of the Special Tokens chain.

                            Token s = t.specialToken;
                            if (s == null)
                            {
                                newToken.next = null;
                                t.specialToken = newToken;
                            }
                            else
                            {
                                while (s.specialToken != null && s.specialToken.kind != BLANK_LINE)
                                {
                                    s = s.specialToken;
                                }

                                newToken.next = s;
                                s.specialToken = newToken;
                            }
                        }
                        else
                        {
                            // Line number found; update it.
                            special.image = Integer.toString(startLineNumber) + " ";
                        }

                        startLineNumber += 10;
                    }
                }
            }
            else if (t.kind == EOL)
            {
                atBOL = true;
            }

            t = t.next;
        }
    }

    public static void main(String[] args) throws Exception
    {
        VBErrHandlerVisitor v = new VBErrHandlerVisitor(SILENT | ADD_HANDLER | APPLY_LINE_NUMBERS, SIMPLE_ERR_HANDLER);

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
            int index = f.getAbsolutePath().lastIndexOf('.');
            if (index < 0)
                throw new IOException(f.getAbsolutePath() + ": No file extension");

            File backup = new File(f.getAbsolutePath().substring(0, index) + ".bak");

            backup.delete();

            if (!f.renameTo(backup))
            {
                throw new IOException("Can not rename " + f);
            }

            parse(backup, f);
        }
    }

    private void parse(File in, File out) throws IOException, ParseException
    {
        VBParser parser = new VBParser(new FileInputStream(in));
        ASTCompilationUnit node = parser.CompilationUnit();
        //node.dump(">");
        //Token t = node.getFirstToken();
        //while (t != null)
        //{
        //  System.out.println(t.debug());
        //  t = t.next;
        //}

        node.jjtAccept(this, null);

        //node.dump("!");

        PrintStream ps = new PrintStream(new FileOutputStream(out));

        ps.println(node.allText(true));
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

