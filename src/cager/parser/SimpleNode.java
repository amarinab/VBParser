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

import java.io.PrintWriter;
import java.io.PrintStream;
import java.util.*;

public class SimpleNode implements Node {
  /**
 * @uml.property  name="parent"
 * @uml.associationEnd  
 */
protected Node parent = null;
  /**
 * @uml.property  name="children"
 * @uml.associationEnd  multiplicity="(0 -1)"
 */
protected Node[] children;
  /**
 * @uml.property  name="id"
 */
protected int id;
  /**
 * @uml.property  name="parser"
 * @uml.associationEnd  
 */
protected VBParser parser;

  public SimpleNode(int i) {
    id = i;
  }

  public SimpleNode(VBParser p, int i) {
    this(i);
    parser = p;
  }

  public void jjtOpen() {
  }

  public void jjtClose() {
  }

  public void jjtSetParent(Node n) { parent = n; }
  public Node jjtGetParent() { return parent; }

  public void jjtAddChild(Node n, int i) {
    if (children == null) {
      children = new Node[i + 1];
    } else if (i >= children.length) {
      Node c[] = new Node[i + 1];
      System.arraycopy(children, 0, c, 0, children.length);
      children = c;
    }
    children[i] = n;
  }

  public Node jjtGetChild(int i) {
    return children[i];
  }

  public int jjtGetNumChildren() {
    return (children == null) ? 0 : children.length;
  }

  /** Accept the visitor. **/
  public Object jjtAccept(VBParserVisitor visitor, Object data) {
    return visitor.visit(this, data);
  }

  /** Accept the visitor. **/
  public Object childrenAccept(VBParserVisitor visitor, Object data) {
    if (children != null) {
      for (int i = 0; i < children.length; ++i) {
        children[i].jjtAccept(visitor, data);
      }
    }
    return data;
  }


	/**
	**	Insert a new child into the list, shifting other elements down.
	**	This is different from jjtAddChild which would overwrite any current child).
	**	These routines can be used to modify the AST after it has been produced, e.g.
	**	to add statements to the end of a procedure etc. The new child[ren] must have
	**	corresponding "Token(s)" created for them which will be written out to any
	**	output file. This method will take care of linking these new tokens into the
	**	token chain.
	**
	**	NB: insertBefore == -1 implies insert at end.
	*/


	/*
	**	This is more difficult than it would first seem. The syntax tree is (obviously)
	**	held as a tree structure, and the token list as a linked list. Each node in the
	**	tree points to a start and end token in the list. However, if we consider all of
	**	the siblings at a particular level, the token ranges they cover will not necessarily
	**	be contiguous. For example given the source.
	**
	**		Public Sub XXXXXX ()
	**			CallThing
	**			CallOtherThing
	**		End Sub
	**
	**	then the parse tree will look like
	**
	**
    **            CompilationUnit
    **                  |
    **            ProcDeclaration
    **            /          \
    **        Name        Statements
    **                    /        \
    **              MethodCall   MethodCall
	**
	**	Each node of this tree covers token ranges as follows:
	**
	**
    **    Public Sub XXXXXX () \n CallThing \n CallOtherThing \n End Sub \n
    **    <--------------------------------------------------------------->       CompilationUnit
    **    <------------------------------------------------------------>          ProcDeclaration
    **               <---->       <---------------------------->                  Name & Statements
    **                            <------->    <------------>                     MethodCall & MethodCall
	**
	**	As can be seen, there may be "gaps" in the token chain at any level. Therefore
	**	we can not assume that (for a particular node)
	**		node.jjtGetChild(i).end.next == node.jjtGetChild(i+1).begin		// W R O N G  ! !
	**	(casts have been omitted from the line above to keep it simple).
	**
	**	Additionally, some child node may actually have a token range that starts BEFORE
	**	the parent's range (e.g. for BinOp nodes). See renumberTokens method below.
	**
	**	There is also an ambiguity. When we insert a new node, do we insert its text
	**	before the following node, or after the preceding node; this will be significant
	**	if there are gap tokens. I have decided to insert before the following token, as
	**	I believe syntactic sugar is more likely to
	*/



	public void insertChild(SimpleNode newChild, int insertBefore) throws ParseException
	{
		insertChild(newChild, insertBefore, null);
	}

	public void insertChild(SimpleNode newChild, int insertBefore, String tokenText) throws ParseException
	{
		if (tokenText != null)
		{
			newChild.begin = Token.newToken(VBParserConstants.ANYTHING_ELSE);
			newChild.begin.image = tokenText;
			newChild.end = newChild.begin;
		}

		SimpleNode c[] = new SimpleNode[1];
		c[0] = newChild;
		insertChild(c, insertBefore);
	}

	public void insertChild(SimpleNode[] newChildren, int insertBefore) throws ParseException
	{
		System.out.println("insertChild");

		if (children == null)
		{
			children = new Node[0];
		}

		if (newChildren[0].begin == null || newChildren[newChildren.length-1].end == null)
		{
			throw new ParseException("No token text for new child");
		}

		if (insertBefore < 0)
			insertBefore = children.length;

		Node c[] = new Node[children.length + newChildren.length];
		if (insertBefore > 0)
			System.arraycopy(children, 0, c, 0, insertBefore);
		System.arraycopy(newChildren, 0, c, insertBefore, newChildren.length);
		if (insertBefore < children.length)
			System.arraycopy(children, insertBefore, c, insertBefore + newChildren.length, children.length - insertBefore);
		children = c;

		/*
		**	That has set up the new JJTree node(s). We now have to link the tokens
		**	into the correct place, as follows:
		**		Find the first token that is to FOLLOW the new token(s).
		**		Update the "next" property of the new token to be this following token.
		**		Change the token which currently points to the "following" token to
		**		point to the new token
		*/

		SimpleNode n;
		Token nextToken;
		int nextSibling = insertBefore + newChildren.length;

		System.out.println("NextSib " + nextSibling + ", insertBefore = " + insertBefore);

		if (nextSibling < children.length)
		{
			/* Insert before last sibling */
			n = (SimpleNode)children[nextSibling];
			if (n.begin == null)
				nextToken = n.end;		/* next sibling consumed no tokens */
			else
				nextToken = n.begin;	/* next sibling did consume tokens */
		}
		else
		{
			/* inserting at the end of the sibling list; insert relative to parent (i.e. this)*/

			n = (SimpleNode)this;
			if (n.begin == null)
				nextToken = n.end;
			else
				nextToken = n.end.next;
		}

		// Now follow the token chain, updaing the correct item.

		insertTokensBefore(newChildren[0].begin, nextToken);
	}

	void insertTokensBefore(Token newTokens, Token insertBefore)
	{
		Token t = findPointerTo(insertBefore);

		if (t == null) throw new Error("insertTokensBefore error: " + newTokens.debug(5) + "//////" + insertBefore.debug(10));

		Token e = newTokens;
		while (e.next != null)
			e = e.next;

		e.next = t.next;

		t.next = newTokens;

		resetBeginEnd();
	}

	private void resetBeginEnd()
	{
		renumberTokens();

		TokenRange range = this.getTokenRange();

		if (range.begin.sequence < begin.sequence)
			begin = range.begin;

		if (range.end.sequence > end.sequence)
			end = range.end;
	}

	public void replaceChild(SimpleNode oldChild, SimpleNode newChild)
	{
		oldChild.dump("REPLACE");
		newChild.dump("BY");

		int child = -1;

		// Find the index of the child to be replaced.
		for (int i = 0; i < children.length; i++)
			if (children[i] == oldChild)
			{
				child = i;
				break;
			}

		if (child < 0) throw new Error("No such child");

		// Update the tree (easy!)
		children[child] = newChild;

		// Now update the token chain (difficult!)

		renumberTokens();

		TokenRange oldRange = oldChild.getTokenRange();
		TokenRange newRange = newChild.getTokenRange();

		System.out.println("Old Range");
		System.out.println(oldRange.debug());
		System.out.println("New Range");
		System.out.println(newRange.debug());

		if (oldRange.begin == null)
		{
			// No old token range.
			System.out.println("No old Range; insert before " + oldRange.end.debug());
			Token t = findPointerTo(oldRange.end);
			t.next = newRange.begin;
			newRange.end.next = oldRange.end;
		}
		else
		{
			System.out.println("Old Range from:");
			System.out.println(oldRange.begin.debug());
			System.out.println("To");
			System.out.println(oldRange.end.debug());
			System.out.println("New Range from:");
			System.out.println(newRange.begin.debug());
			System.out.println("To");
			System.out.println(newRange.end.debug());
			Token t = findPointerTo(oldRange.begin);
			t.next = newRange.begin;
			newRange.end.next = oldRange.end.next;
		}
	}

	public Node findRoot()
	{
		Node node = this;
		while (node.jjtGetParent() != null)
		{
			node = node.jjtGetParent();
		}

		return node;
	}

	private Token findPointerTo(Token target)
	{
		SimpleNode root = (SimpleNode)findRoot();

		for (Token t = root.begin; t != null; t = t.next)
		{
			if (t.next == target)
			{
				return t;
			}
		}

		return null;
	}


/*
	private SimpleNode findNextNode(SimpleNode node, int start, boolean ignoreEmpty)
	{
		System.out.println("findNextNode " + start + " " + node.toString());
		for (int i = start; i < node.jjtGetNumChildren(); i++)
		{
			System.out.println("Child " + i + ": " + ((SimpleNode)node.jjtGetChild(i)).begin);
			if (!ignoreEmpty || ((SimpleNode)node.jjtGetChild(i)).begin != null)
			{
				System.out.println("Return child " + i);
				return (SimpleNode)(node.jjtGetChild(i));
			}
		}

		// Not found in "node" -- check its parents.

		int myIndex = -1;
		SimpleNode parent = (SimpleNode)node.jjtGetParent();
		if (parent == null)
			return null;

		for (int i = 0; i < parent.jjtGetNumChildren(); i++)
		{
			if (parent.jjtGetChild(i) == node)
			{
				myIndex = i;
				break;
			}
		}

		if (myIndex == -1)
			throw new Error("Coding Error");

		System.out.println("Trying parent at pos " + (myIndex + 1));

		return findNextNode(parent, myIndex + 1, ignoreEmpty);
	}


	private Token findFollowing(SimpleNode node, int index)
	{
		System.out.println("Search node " + VBParserTreeConstants.jjtNodeName[node.id] + " at index " + index);
		for (int i = index + 1; i < node.jjtGetNumChildren(); i++)
		{
			if (((SimpleNode)node.jjtGetChild(i)).begin != null)
			{
				System.out.println("Found at " + i);
				return ((SimpleNode)node.jjtGetChild(i)).begin;
			}
		}

		// Insert at end of this node (or end of the parent of this node...)

		while (node.end == null)
		{
			if (node.jjtGetParent() == null)
				return null;

			System.out.println("node.end == null. Trying " + VBParserTreeConstants.jjtNodeName[((SimpleNode)node.jjtGetParent()).id]);

			node = (SimpleNode)(node.jjtGetParent());
		}

		System.out.println("Setting following to " + node.end.next.sequence + "\"" + node.end.next.image + "\"");

		return node.end.next;
	}
	*/

/*
	private Token findFollowing(SimpleNode node, int index)
	{
		System.out.println("Search node " + VBParserTreeConstants.jjtNodeName[node.id] + " at index " + index);
		for (int i = index + 1; i < node.jjtGetNumChildren(); i++)
		{
			if (((SimpleNode)node.jjtGetChild(i)).begin != null)
				return ((SimpleNode)node.jjtGetChild(i)).begin;
		}

		// Try parent (recursively).

		SimpleNode p = (SimpleNode)(node.jjtGetParent());

		if (p == null)
		{
			// Top-level node.

			return null;
		}

		// What index in the parent's child collection is this node?

		int myIndex = 0;
		for (myIndex = 0; myIndex < p.jjtGetNumChildren(); myIndex++)
		{
			if (p.jjtGetChild(myIndex) == node)
			{
				break;
			}
		}

		return findFollowing(p, myIndex);
	}
*/



  /* You can override these two methods in subclasses of SimpleNode to
     customize the way the node appears when the tree is dumped.  If
     your output uses more than one line you should override
     toString(String), otherwise overriding toString() is probably all
     you need to do. */

  public String toString()
  {
	  StringBuffer buff = new StringBuffer();

	  if (begin != null)
	  {
		buff.append(Integer.toString(begin.beginLine));
		buff.append(" ");
	  }

	  buff.append(VBParserTreeConstants.jjtNodeName[id]);
	  buff.append(":  ");
	  if (begin != null)
	  {
	  	buff.append(begin.image);
	  	if (begin.next != null && begin != end) buff.append("...");
      }

	  return buff.toString();
  }

  public String toString(String prefix) { return prefix + toString(); }

  /* Override this method if you want to customize how the node dumps
     out its children. */

	public void dump(PrintStream p, String prefix)
	{
		p.println(toString(prefix));
		if (children != null) {
			for (int i = 0; i < children.length; ++i) {
				SimpleNode n = (SimpleNode)children[i];
				if (n != null) {
					n.dump(p, prefix + "    ");
				}
			}
		}
	}

	public void dump(PrintWriter p, String prefix)
	{
		p.println(toString(prefix));
		if (children != null) {
			for (int i = 0; i < children.length; ++i) {
				SimpleNode n = (SimpleNode)children[i];
				if (n != null) {
					n.dump(p, prefix + "    ");
				}
			}
		}
	}

	public void dump(String prefix)
	{
		dump(System.out, prefix);
	}

	static public String getTokenText(Token t)
	{
		return getTokenText(t, false);
	}

	static public String getTokenText(Token t, boolean includeSpecialText)
	{
		Token tt = t.specialToken;

		if (!includeSpecialText || tt == null)
			return t.image;

		StringBuffer buff = new StringBuffer();

		// Find the start of the special token chain.
		while (tt.specialToken != null)
			tt = tt.specialToken;
		// Now follow the chain forwards, collecting text
		while (tt != null)
		{
			buff.append(tt.image);
			tt = tt.next;
		}

		buff.append(t.image);

		return buff.toString();
	}

	private static final String spaces = "                                                           ";

	public String allText()
	{
		return allText(false);
	}

	public String allText(boolean includeSpecialText)
	{
		Token t = begin;
		StringBuffer buff = new StringBuffer();

		// TODO - this might mean we lose any comment at the end of a file without a \n.
		// Very unlikely...
		while (t != null)
		{
			if (t.kind != VBParserConstants.EOF)
				buff.append(getTokenText(t, includeSpecialText));

			if (t == end)
				break;

			t = t.next;
		}

		return buff.toString();
	}

/*
	public String allText(boolean includeSpecialText, int depth) throws ParseException
	{
//		System.out.println(spaces.substring(0, depth*2) +
//			VBParserTreeConstants.jjtNodeName[id] +
//			" " + begin.sequence + " <" + begin.image + "> " + end.sequence + " <" + end.image);

		Token t = begin;
		StringBuffer buff = new StringBuffer();
		int childIndex = 0;

		while (t != null)
		{
			if (children != null && childIndex < children.length)
			{
				SimpleNode child = (SimpleNode)(children[childIndex]);

				// No tokens for this child?

				if (child.begin == null)
				{
					for ( ; childIndex < children.length; childIndex++)
					{
						SimpleNode next = (SimpleNode)(children[childIndex]);
						if (next.begin != child.begin)
							break;
					}

					continue;
				}

				// Have we reached the start of this child's tokens?

				if (child.begin == t)
				{
//					System.out.println(spaces.substring(0, depth*2) + "Call child");
					buff.append(child.allText(includeSpecialText, depth + 1));

					if (child.end == end)
					{
						return buff.toString();
					}

					t = child.end.next;

//					System.out.println(spaces.substring(0, depth*2) + "END of child");

					for ( ; childIndex < children.length; childIndex++)
					{
						SimpleNode next = (SimpleNode)(children[childIndex]);
						if (next.begin != child.begin)
							break;
					}

//					System.out.println(spaces.substring(0, depth*2) + "Next Child " + childIndex);
					if (childIndex < children.length)
					{
						child = (SimpleNode)(children[childIndex]);
//						System.out.println(spaces.substring(0, depth*2) + "           " +
//							VBParserTreeConstants.jjtNodeName[child.id] +
//							" " + child.begin.sequence + " <" + child.begin.image + "> " + child.end.sequence + " <" + child.end.image);
					}
					else
					{
//						System.out.println(spaces.substring(0, depth*2) + "           " + "No more children");
					}

					continue;
				}
			}

			buff.append(getTokenText(t, includeSpecialText));
//			System.out.println(spaces.substring(0, depth*2) + "Appended <" + getTokenText(t, false) + ">");

			if (t == end)
			{
//				System.out.println(spaces.substring(0, depth*2) + "END of this node at " + t.sequence + " " + t.image);
				break;
			}

			t = t.next;
		}

//		System.out.println(spaces.substring(0, depth*2) + "return " + VBParserTreeConstants.jjtNodeName[id] + ": <" + buff.toString() + ">");
		return buff.toString();
	}
*/

	public SimpleNode findNode(int nodeId)
	{
		return findNode(nodeId, null);
	}

	public SimpleNode findNode(int nodeId, String name)
	{
		SimpleNode ret;

		if ((nodeId == id) && (name == null || (begin != null && name.equalsIgnoreCase(begin.image))))
			return this;

		if (children == null)
			return null;

		for (int i = 0; i < children.length; i++)
		{
			SimpleNode n = (SimpleNode)(children[i]);
			ret = n.findNode(nodeId, name);

			if (ret != null)
				return ret;
		}

		return null;
	}

	public SimpleNode[] findNodes(int nodeId)
	{
		return findNodes(nodeId, null);
	}

	public SimpleNode[] findNodes(int nodeId, String name)
	{
		Vector v = new Vector();
		findNodes(nodeId, name, v);
		return (SimpleNode[])v.toArray(new SimpleNode[v.size()]);
	}

	public void findNodes(int nodeId, String name, Vector v)
	{
		//System.out.println("Node Id " + id + " (target " + nodeId + ")");

		if ((nodeId == id) && (name == null || (begin != null && name.equalsIgnoreCase(begin.image))))
			v.add(this);

		if (children == null)
			return;

		for (int i = 0; i < children.length; i++)
		{
			SimpleNode n = (SimpleNode)(children[i]);
			n.findNodes(nodeId, name, v);
		}
	}


	/*
	**	Each SimpleNode has a range of tokens (the 'begin' and 'end' variables).
	**	However parent / child nodes do not always have the ranges that would be
	**	expected. For example a child node may have a range that starts before
	**	the parent's range (e.g. for BinOp productions:
			void ImpExpression() :
			{}
			{
				EqvExpression()
				(
					( <IMP> { jjtThis.op = getToken(0).image; } EqvExpression() ) #BinOp(2)
				)*
			}
	**	We therefore need a convenient way to determine token ranges, and in particular, we
	**	need an efficient method to determine if one token is ealier in the chain than another.
	**	To do this we look at the tokens' sequence numbers.
	**
	**	This method is provided as tokens may be inserted or deleted from the token chain,
	*/
	public void renumberTokens()
	{
		SimpleNode n = this;
		while (n.jjtGetParent() != null)
			n = (SimpleNode)n.jjtGetParent();

		int seq = 1;
		for (Token t = n.begin; t != null; t = t.next)
			t.sequence = seq++;

//		System.out.println("Renumber debug:");
//		for (Token t = n.begin; t != null; t = t.next)
//			System.out.println(t.debug());
	}

	public TokenRange getTokenRange()
	{
		TokenRange r = new TokenRange();

		recurseTokenRange(this, r);

		return r;
	}

	private void recurseTokenRange(SimpleNode n, TokenRange r)
	{
		if (r.begin == null)
			r.begin = n.begin;
		else if (n.begin != null && n.begin.sequence < r.begin.sequence)
			r.begin = n.begin;

		if (r.end == null)
			r.end = n.end;
		else if (n.end != null && n.end.sequence > r.end.sequence)
			r.end = n.end;

		if (n.children != null)
		{
			for (int i = 0; i < n.children.length; i++)
			{
				SimpleNode child = (SimpleNode)n.children[i];

				recurseTokenRange(child, r);
			}
		}
	}



  /**
 * @uml.property  name="begin"
 * @uml.associationEnd  
 */
protected Token begin;
/**
 * @uml.property  name="end"
 * @uml.associationEnd  
 */
protected Token end;

  public void setFirstToken(Token t) { begin = t;}
  public void setLastToken(Token t) { end = t;}

  public Token getFirstToken() { return begin;}
  public Token getLastToken() { return end;}

  /**
 * @uml.property  name="parseException"
 * @uml.associationEnd  
 */
public ParseException parseException = null;

  /**
 * @uml.property  name="name"
 */
protected String name;
  /**
 * @return
 * @uml.property  name="name"
 */
public String getName()
  {
	  return name;
  }
  /**
 * @param name
 * @uml.property  name="name"
 */
public void setName(String name)
  {
	  this.name = name;
  }

  public void process(PrintWriter ostr)
  {
  }

  protected void print(Token t, PrintWriter ostr)
  {
	  Token tt = t.specialToken;

	  if (tt != null)
	  {
		  while (tt.specialToken != null) tt = tt.specialToken;

		  while (tt != null) {
			  ostr.print(tt.image);
			  tt = tt.next;
		  }
	  }

	  ostr.print(t.image);
  }

  /**
 * @author  Tony
 */
private static class TokenRange
  {
	  public TokenRange()
	  {
	  }

	  public TokenRange(Token begin, Token end)
	  {
		  this.begin = begin;
		  this.end = end;
	  }

	  /**
	 * @uml.property  name="begin"
	 * @uml.associationEnd  
	 */
	private Token begin;
	  /**
	 * @uml.property  name="end"
	 * @uml.associationEnd  
	 */
	private Token end;

	  public String debug()
	  {
		  if (begin == null)
		  {
			  return "TokenRange: start <null> end " + (end == null ? "<null>" : end.debug());
		  }
		  else
		  {
			  StringBuffer sb = new StringBuffer();
			  sb.append("TokenRange:");
			  for (Token t = begin; ; t = t.next)
			  {
				  sb.append(t.debug());
				  if (t == end)
				  	break;
			  }
			  return new String(sb);
		  }
	  }
  }
}

