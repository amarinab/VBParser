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
/* Generated By:JavaCC: Do not edit this line. Token.java Version 3.0 */

package cager.parser;

/**
 * Describes the input token stream.
 */

public class Token {

	static long nextSequence = 0;

	public Token()
	{
		sequence = nextSequence++;
	}

	public Token(int ofKind)
	{
		this();
		kind = ofKind;
	}

  /**
 * An integer that describes the kind of this token.  This numbering system is determined by JavaCCParser, and a table of these numbers is stored in the file ...Constants.java.
 * @uml.property  name="kind"
 */
  public int kind;

  /**
 * beginLine and beginColumn describe the position of the first character of this token; endLine and endColumn describe the position of the last character of this token.
 * @uml.property  name="beginLine"
 */
  public int beginLine;

/**
 * beginLine and beginColumn describe the position of the first character of this token; endLine and endColumn describe the position of the last character of this token.
 * @uml.property  name="beginColumn"
 */
public int beginColumn;

/**
 * beginLine and beginColumn describe the position of the first character of this token; endLine and endColumn describe the position of the last character of this token.
 * @uml.property  name="endLine"
 */
public int endLine;

/**
 * beginLine and beginColumn describe the position of the first character of this token; endLine and endColumn describe the position of the last character of this token.
 * @uml.property  name="endColumn"
 */
public int endColumn;

  /**
 * The string image of the token.
 * @uml.property  name="image"
 */
  public String image;

  /**
 * A reference to the next regular (non-special) token from the input stream.  If this is the last token from the input stream, or if the token manager has not read tokens beyond this one, this field is set to null.  This is true only if this token is also a regular token.  Otherwise, see below for a description of the contents of this field.
 * @uml.property  name="next"
 * @uml.associationEnd  inverse="specialToken:cager.parser.Token"
 */
  public Token next;

  /**
 * This field is used to access special tokens that occur prior to this token, but after the immediately preceding regular (non-special) token. If there are no such special tokens, this field is set to null. When there are more than one such special token, this field refers to the last of these special tokens, which in turn refers to the next previous special token through its specialToken field, and so on until the first special token (whose specialToken field is null). The next fields of special tokens refer to other special tokens that immediately follow it (without an intervening regular token).  If there is no such token, this field is null.
 * @uml.property  name="specialToken"
 * @uml.associationEnd  inverse="next:cager.parser.Token"
 */
  public Token specialToken;

  /**
   * Returns the image.
   */
  public final String toString()
  {
     return image;
  }

  public String debug()
  {
	  return debug(1);
  }

  public String debug(int numTokens)
  {
	  StringBuffer sb = new StringBuffer();

	  Token t = this;

	  for (int i = 1; i <= numTokens; i++)
	  {
		  if (t == null)
		  	break;

		  if (t.specialToken != null)
		  {
			sb.append("Special Tokens:\n");
			Token s = t.specialToken;
			while (s.specialToken != null)
				s = s.specialToken;
			while (s != null)
			{
				sb.append("    " + VBParserConstants.tokenImage[s.kind] + "  <" + s.image + ">\n");
				s = s.next;
			}
		  }

		  sb.append("[" + t.sequence + "] " + VBParserConstants.tokenImage[t.kind] + "  <" + t.image + ">\n");

		  t = t.next;
	  }

	  return sb.toString();
  }

  /**
 * @uml.property  name="sequence"
 */
public long sequence;

  /**
 * @uml.property  name="currentLineNumberLabel"
 */
public String currentLineNumberLabel = null;

  /**
   * Returns a new Token object, by default. However, if you want, you
   * can create and return subclass objects based on the value of ofKind.
   * Simply add the cases to the switch for all those special cases.
   * For example, if you have a subclass of Token called IDToken that
   * you want to create if ofKind is ID, simlpy add something like :
   *
   *    case MyParserConstants.ID : return new IDToken();
   *
   * to the following switch statement. Then you can cast matchedToken
   * variable to the appropriate type and use it in your lexical actions.
   */
  public static final Token newToken(int ofKind)
  {
     switch(ofKind)
     {
       default : return new Token(ofKind);
     }
  }

}