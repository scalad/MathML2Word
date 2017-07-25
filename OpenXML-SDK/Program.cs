using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Xml.Xsl;
using DocumentFormat.OpenXml;
using System.Xml;
using System.IO;
using Microsoft.Office.Interop.Word;
using DocumentFormat;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXML_SDK
{
    class Program
    {
           
        public static void MathML2Word()
        {
            XslCompiledTransform xslTransform = new XslCompiledTransform();
            xslTransform.Load(@"C:\Program Files (x86)\Microsoft Office\Office14\MML2OMML.xsl");

            // Load the file containing your MathML presentation markup.
            using (XmlReader reader = XmlReader.Create(File.Open("mathML1.xml", FileMode.Open)))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlWriterSettings settings = xslTransform.OutputSettings.Clone();

                    // Configure xml writer to omit xml declaration.
                    settings.ConformanceLevel = ConformanceLevel.Fragment;
                    settings.OmitXmlDeclaration = true;
                    XmlWriter xw = XmlWriter.Create(ms, settings);
                    // Transform our MathML to OfficeMathML
                    xslTransform.Transform(reader, xw);
                    ms.Seek(0, SeekOrigin.Begin);
                    StreamReader sr = new StreamReader(ms, Encoding.UTF8);
                    string officeML = sr.ReadToEnd();
                    Console.Out.WriteLine(officeML);
                    // Create a OfficeMath instance from the OfficeMathML xml.
                    DocumentFormat.OpenXml.Math.OfficeMath om = new DocumentFormat.OpenXml.Math.OfficeMath(officeML);


                    //创建Word文档(Microsoft.Office.Interop.Word)  
                    Microsoft.Office.Interop.Word._Application WordApp = new Application();
                    WordApp.Visible = true;
                    using (WordprocessingDocument package = WordprocessingDocument.Create("template.docx", WordprocessingDocumentType.Document))
                    {
                        // Add a new main document part. 
                        package.AddMainDocumentPart();

                        // Create the Document DOM. 
                        package.MainDocumentPart.Document =
                          new DocumentFormat.OpenXml.Wordprocessing.Document(
                            new Body(
                              new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                new Run(
                                  new Text("Hello World!")))));
                         
                        // Save changes to the main document part. 
                        package.MainDocumentPart.Document.Save(); 
                    }
                    
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("template.docx", true))
                    {
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph par =
                          wordDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();

                        foreach (var currentRun in om.Descendants<DocumentFormat.OpenXml.Math.Run>())
                        {
                            // Add font information to every run.
                            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties2 =
                              new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                            currentRun.InsertAt(runProperties2, 0);
                        }
                        par.Append(om);
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            MathML2Word();
            Console.ReadLine();
        }
    }
}
