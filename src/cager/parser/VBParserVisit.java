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

public class VBParserVisit implements VBParserVisitor
{
  public Object visit(SimpleNode node, Object data){ node.childrenAccept(this, data); return data; }

  public Object visit(ASTCompilationUnit node, Object data)
  {
	  start();
	  node.childrenAccept(this, data);
	  end();
	  return data;
  }

  public Object visit(ASTProcDeclaration node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTStatements node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTStatement node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTAssignment node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTMethodCall node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTOption node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTDeclaration node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTReDim node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTEventDeclaration node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTDeclare node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTConstDeclaration node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTTypeDeclaration node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTEnumDeclaration node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTSetStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTOnError node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTIfStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTImplements node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTDoWhileStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTDoCondition node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTWhileWendStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTForEachStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTForStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTWithStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTCaseStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTArgList node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTExitStatement node, Object data){ return visit((ASTStatement)node, data); }
  public Object visit(ASTName node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTExprTest node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTExpression node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTBinOp node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTUnaryOp node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTPrimaryExpression node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTLiteral node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTLabel node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTTypeName node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTFormalParamList node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTVarDim node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTArgSpec node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTVarDecl node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTUnrecognisedStatement node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTArguments node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTCaseExpr node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTFormItem node, Object data){ node.childrenAccept(this, data); return data; }
  public Object visit(ASTParamSpec node, Object data){ node.childrenAccept(this, data); return data; }

  public void start() { }
  public void end() { }

}
