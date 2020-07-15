using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace StyleRenamer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Status.Text = "";
        }

        private void GetInputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog theDialogue = new OpenFileDialog();
            theDialogue.Filter = "MS Word|*.docx";
            if ((bool)theDialogue.ShowDialog())
                FileToProcess.Text = theDialogue.FileName;
        }

        private void ProcessStyles_Click(object sender, RoutedEventArgs e)
        {
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            Status.Text = "";
            int Counter = 0;
            string OutFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(FileToProcess.Text), System.IO.Path.GetFileNameWithoutExtension(FileToProcess.Text) + ".xml");
            using (WordprocessingDocument doc =
            WordprocessingDocument.Open(FileToProcess.Text, true))
            {
                // Get the Styles part for this document.
                StyleDefinitionsPart part =
                doc.MainDocumentPart.StyleDefinitionsPart;
                foreach (DocumentFormat.OpenXml.OpenXmlElement theElement in part.Styles.Elements())
             
                    try
                    {
                        OpenXmlAttribute CustomStyle = theElement.GetAttribute("customStyle", w);
                        if (CustomStyle.Value == "1")
                        {
                             foreach (OpenXmlElement theChildElement in theElement.Elements())
                             {
                                 if (theChildElement.LocalName == "name")
                                 {
                                     OpenXmlAttribute theName = theChildElement.GetAttribute("val", w);
                                     string tmpString = theName.Value;
                                     Counter++;
                                     theName.Value += " " + Counter.ToString();
                                 }
                             }

                        }
                    }
                    catch
                    { }
                doc.Close();
                
                Status.Text = Counter.ToString() + " styles processed";

                }
            }
        
        }
    }

                    
        






            //using (Package wdPackage = Package.Open(FileToProcess.Text, FileMode.Open, FileAccess.Read))
            //{
            //    PackageRelationship docPackageRelationship = wdPackage.GetRelationshipsByType(documentRelationshipType).FirstOrDefault();
            //    if (docPackageRelationship != null)
            //    {
            //        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), docPackageRelationship.TargetUri);
            //        PackagePart documentPart = wdPackage.GetPart(documentUri);

            //        //  Load the document XML in the part into an XDocument instance.
            //        xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

            //        //  Find the styles part. There will only be one.
            //        PackageRelationship styleRelation = documentPart.GetRelationshipsByType(stylesRelationshipType).FirstOrDefault();
            //        if (styleRelation != null)
            //        {
            //            Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
            //            PackagePart stylePart = wdPackage.GetPart(styleUri);

            //            //  Load the style XML in the part into an XDocument instance.
            //            stylesDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
            //            int Counter = 0;
            //            foreach (XElement style in stylesDoc.Root.Elements(w + "style"))
            //            {
            //                if (style.Attribute(w + "customStyle") != null)
            //                {
            //                    if (style.Attribute(w + "customStyle").Value == "1")
            //                    {
            //                        XElement theName = style.Element(w + "name");
            //                        if (theName != null)
            //                        {
            //                            Counter++;
            //                            string theNameValue = theName.Attribute(w + "val").Value;
            //                            theName.Attribute(w + "val").SetValue(theNameValue + " " + Counter.ToString());
            //                            theNameValue = theName.Attribute(w + "val").Value;
            //                        }

            //                    }

            //                }
            //            }
            //             using (XmlWriter xw = XmlWriter.Create(OutFile))
            //            {
            //                xDoc.Add(stylesDoc.Root);
            //                xDoc.Save(xw);
                  
            //            }
            //           //Doc.Save(OutFile); // Save the document


        //            }
        //            wdPackage.Close();
        //            Status.Text = "Finished";
        //        }
                
 
        //    }


        //}
       
