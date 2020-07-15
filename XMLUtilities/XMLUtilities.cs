using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.IO;
using WordApp = Microsoft.Office.Interop.Word._Application;
using WordRoot = Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word._Document;

/*
 * This contains a number of utilities to help manipulate Word documents in XML Form
 * 
 * Last modified by Stephen Palmstrom 2 February 2015
 */
namespace XMLUtilities
{
    public class XMLUtility
    {
        const string wordmlNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
        const string wordmlxNamespace = "http://schemas.microsoft.com/office/word/2003/auxHint";

        public bool IsInitialised = false;
        /*
         * A number of dictionaries to hold information on the document
         */
        public Dictionary<string, string> theStyleDictionary = new Dictionary<string, string>(10); // to hold all defined styles
        public Dictionary<string, string> theDefaultStyleDictionary = new Dictionary<string, string>(5); // to hold all default styles
        private XmlNamespaceManager nsManager = null;
        private XmlDocument theXMLDocument;
        private XmlNode theRoot = null;

        public XMLUtility(WordApp wrdApp, Document theWordDocument)
        {
            /*
             * We initialise various things, load styles, fonts etc. for future use.
             */
            IsInitialised = true;
            theXMLDocument = new XmlDocument();
            theWordDocument.Select();
            theXMLDocument.LoadXml(wrdApp.Selection.get_XML(false));
            theRoot = theXMLDocument.DocumentElement;  // the root node
            if (theXMLDocument != null)
            {
                nsManager = new XmlNamespaceManager(theXMLDocument.NameTable);
                nsManager.AddNamespace("wx", wordmlxNamespace);
                nsManager.AddNamespace("w", wordmlNamespace);
                // If successful add the root to the dictionary
            }
            else
            {
                // We failed

            }

            GetStylesInUse(theRoot, nsManager, theStyleDictionary);  // Load the styles

        }



        ~XMLUtility()
        {
            // the destructor
            theStyleDictionary = null;
            theDefaultStyleDictionary = null;
        }


        private void GetStylesInUse(XmlNode theRoot, XmlNamespaceManager nsManager, Dictionary<string, string> theStyleDictionary)
        {                // Load a list of current styles and their fonts
            XmlNodeList theNodeList = theRoot.SelectNodes(@"//w:styles/w:style", nsManager);
            theStyleDictionary.Clear();  // Empty the style dictionary
            // First look for the styles that have fonts
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    string theFontName = theFont.Attributes["wx:val"].Value;
                    theStyleDictionary.Add(theStyleID, theFontName);
                }
            }
            // Now look for the default fonts - we do this as a second pass in case they don't appear first
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    if (theStyle.Attributes.Count == 3)
                    {
                        // check to see if this is the default.
                        bool IsDefault = false;
                        try
                        {
                            IsDefault = (theStyle.Attributes[@"w:default"].Value == "on" /*&& theStyle.Attributes[@"w:type"].Value == "paragraph"*/);
                        }
                        catch
                        {
                        }
                        if (IsDefault)
                        {


                            string theDefaultStyle = theStyle.Attributes[@"w:styleId"].Value;
                            // We have found a default style so we look up its font and add to the nominal styles.
                            switch (theStyle.Attributes[@"w:type"].Value)
                            {
                                case "paragraph":
                                    theStyleDictionary["DefaultParagraphFont"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "table":
                                    theStyleDictionary["Default Table"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "character":
                                    theStyleDictionary["Default Character"] = theStyleDictionary[theDefaultStyle];
                                    break;

                            }
                        }
                    }
                }
            }
            // Now the styles that don't have fonts- we have to get the font of the style on which they are based.
            // but first we load those that aren't based on a style 
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                XmlNode theBasedOnStyle = theStyle.SelectSingleNode("w:basedOn", nsManager);
                if (theFont == null && theBasedOnStyle == null)
                {
                    // Use the default paragraph font
                    theStyleDictionary[theStyleID] = theStyleDictionary["DefaultParagraphFont"];
                }
            }
            // Now look at the styles that don't have fonts but are based on other styles, and give them the fonts from the styles on which they were based
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                XmlNode theBasedOnStyle = theStyle.SelectSingleNode("w:basedOn", nsManager);
                if (theFont == null && theBasedOnStyle != null)
                {
                    // Use the default paragraph font
                    string theBasedOnStyleID = theBasedOnStyle.Attributes["w:val"].Value;
                    theStyleDictionary[theStyleID] = theStyleDictionary[theBasedOnStyleID];
                }
            }

        }
        public List<RichText> GetText(WordRoot.Paragraph theParagraph)
        {
            /*
             * We get a text string and its corresponding font.
             */
            XmlDocument theXMLDoc = new XmlDocument();
            theXMLDoc.LoadXml(theParagraph.Range.WordOpenXML);
            XmlNode theXMLNode = theXMLDoc.DocumentElement;
            XmlNodeList theNodeList =  theXMLNode.SelectNodes(@"//w:body//w:p", nsManager);  // Find the paragraphs
            string theParagraphFont = null;
            string TextString = null;
            List<RichText> theRichText = new List<RichText>();
            foreach (XmlNode theParagraphData in theNodeList)
            {
                string theParagraphStyleID = XmlLookup(theParagraphData, "w:pPr/w:pStyle", nsManager, "w:val", "DefaultParagraphFont");
                RichText TextElement = new RichText()
                {
                    Font = null
                };
                if (theStyleDictionary.Keys.Contains(theParagraphStyleID))
                {
                    TextElement.Font = theStyleDictionary[theParagraphStyleID];
              }
                else
                {
                    TextElement.Font = GetDefaultFont(theStyleDictionary, theParagraphData);
                }
                //}
                theParagraphFont = TextElement.Font;  // Remember the paragraph font for the end of line.
                XmlNodeList theRanges = theParagraphData.SelectNodes("w:r", nsManager);
                TextElement.Text = "";
                /*
                 * We go through the document a range at a time.  If we find a symbol whose font is the same as that of an existing range
                 * we concatenate the symbol to that range.
                 */
                string OldFontName   = "";
                foreach (XmlNode theRangeData in theRanges)
                {
                    XmlNode theSymbol = theRangeData.SelectSingleNode("w:sym", nsManager);
                    if (theSymbol != null)
                    {
                        // we have a symbol
                        TextElement.Font = theSymbol.Attributes["w:font"].Value;
                        string theSymbolValue = theSymbol.Attributes["w:char"].Value;
                        char theChar = Convert.ToChar(Convert.ToUInt16(theSymbolValue, 16));  // get the character number
                        if (TextElement.Font == OldFontName)
                        {
                            // Concatenate the text string
                            TextElement.Text += Convert.ToString(theChar); // make it a string concatenating it with previous symbols.
                        }
                        else
                        {
                            OldFontName = TextElement.Font;
                            TextString = Convert.ToString(theChar); // make it a string concatenating it with previous symbols. 
                        }

                    }
                    else
                    {

                        // See if there is a font defined in the range and use that
                        TextElement.Font = XmlLookup(theRangeData, "w:rPr/wx:font", nsManager, "wx:val", "");
                        if (TextElement.Font == "")
                        {
                            string theStyleID = XmlLookup(theRangeData, "w:rPr/w:rStyle", nsManager, "w:val", "");
                            if (theStyleID != "" && theStyleDictionary.Keys.Contains(theStyleID))
                            {
                                // If we have no style nor do we have a font for the style, we do nothing
                                // Otherwise we get the font name for the style.
                                TextElement.Font = theStyleDictionary[theStyleID];
                            }
                            else
                            {
                                TextElement.Font = theParagraphFont; // we pick up the paragraph font
                            }
                        }

                        // Look for text
                        XmlNode theText = theRangeData.SelectSingleNode("w:t", nsManager);
                        if (theText != null)
                        {
                            if (TextElement.Font == OldFontName)
                            {
                                TextString += theText.InnerText;
                            }
                            else
                            {
                                OldFontName = TextElement.Font;
                                TextString = theText.InnerText;
                            }
                        }

                    }
                }
                TextElement.Text = TextString;
                theRichText.Add(TextElement);

            }

            return theRichText;
        }

        private string XmlLookup(XmlNode theNode, string theSearchPath, XmlNamespaceManager nsManager, string theValueID, string InputString = "")
        {
            /*
             * This looks up something in Xml and updates the input string with the returned value.  Otherwise
             * it returns the input string.  The idea is to update some data with new information
             */
            XmlNode theChildNode = theNode.SelectSingleNode(theSearchPath, nsManager);
            if (theChildNode == null)
            {
                // We didn't find anything, so return the input string
                return InputString;
            }
            else
            {
                try
                {
                    string tmpString = theChildNode.Attributes[theValueID].Value;
                    return tmpString;
                }
                catch (Exception Ex)
                {
                    // Something went wrong
                    string theError = Ex.Message;
                    return InputString;
                }
            }
        }
        private string GetDefaultFont(Dictionary<string, string> theStyleDictionary, XmlNode theNode)
        {
            string NodePath = GetNodePath(theNode, "");
            string theDefaultID = "DefaultParagraphFont";  // we assume a normal paragraph
            if (NodePath.Contains("w:tbl"))
            {
                // we have a table
                theDefaultID = "Default Table";
            }

            return theStyleDictionary[theDefaultID];
        }
        private string GetNodePath(XmlNode theNode, string InputType)
        {
            // Iteratively walk up the nodes.
            XmlNode theParent = theNode.ParentNode;
            if (theParent != null)
            {
                string tmpString = theParent.Name + "/" + InputType;
                GetNodePath(theParent, tmpString);
                return tmpString;
            }
            else
            {
                return InputType;
            }
        }


    }
    public class RichText
    {
        public string Font;
        public string Text;
    }
}
