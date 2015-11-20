using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeWhitener
{
    /// <summary>
    /// Provides methods to remove personal informations from Office files.
    /// </summary>
    /// <seealso cref="https://msdn.microsoft.com/en-us/library/bb739835(v=office.12).aspx#ManipulatingWord2007OpenXMLFiles_RetrieveaDocumentCoreProperty"/>
    /// <seealso cref="https://sites.google.com/site/compositiosystemae/home/vbaworld/office10/xlsx"/>
    /// <seealso cref="http://blogs.msdn.com/b/ericwhite/archive/2008/07/13/using-the-open-xml-sdk-and-linq-to-xml-to-remove-personal-information-from-an-open-xml-wordprocessing-document.aspx"/>
    static class OfficeDocumentWhitener
    {
        static string[] _AvailableExtensions = new string[] { "docx", "docm", "xlsx", "xlsm", "pptx", "pptm" };

        /// <summary>
        /// Extensions that can be supported to remove personal informations
        /// </summary>
        public static string[] AvailableExtensions { get { return _AvailableExtensions; } }

        /// <summary>
        /// Removes personal informations from a Word document, Excel spreadsheet or PowerPoint presentation according to the extension.
        /// </summary>
        /// <param name="filepath"></param>
        public static void RemovePersonalInfo(string filepath)
        {
            if (!File.Exists(filepath)) throw new FileNotFoundException("File Not Found", filepath);

            switch (Path.GetExtension(filepath).Replace(".", ""))
            {
                case "docx":
                case "docm":
                    using (var doc = WordprocessingDocument.Open(filepath, true))
                    {
                        OfficeDocumentWhitener.RemovePersonalInfo(doc);
                    }
                    break;
                case "xlsx":
                case "xlsm":
                    using (var doc = SpreadsheetDocument.Open(filepath, true))
                    {
                        OfficeDocumentWhitener.RemovePersonalInfo(doc);
                    }
                    break;
                case "pptx":
                case "pptm":
                    using (var doc = PresentationDocument.Open(filepath, true))
                    {
                        OfficeDocumentWhitener.RemovePersonalInfo(doc);
                    }
                    break;
                case "doc":
                case "xls":
                case "ppt":
                    throw new NotSupportedException("Old office format is not supported.");
                default:
                    throw new NotSupportedException("This format is not supported.");
            }
        }

        /// <summary>
        /// Removes personal informations from a Word document.
        /// </summary>
        /// <param name="document"></param>
        public static void RemovePersonalInfo(WordprocessingDocument document)
        {
            RemovePersonalInfo(document.ExtendedFilePropertiesPart, document.CoreFilePropertiesPart);

            var documentSettingsPart = document.MainDocumentPart.DocumentSettingsPart;
            XNamespace ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // add w:removePersonalInformation, w:removeDateAndTime to /word/settings.xml
            XDocument documentSettings = documentSettingsPart.GetXDocument();

            //Debug.WriteLine(documentSettings.ToString());

            // add the new elements in the right position.  Add them after the following three elements
            // (which may or may not exist in the xml document).
            XElement settings = documentSettings.Root;
            XElement lastOfTop3 = settings.Elements()
                .Where(e => e.Name == ns + "writeProtection" ||
                    e.Name == ns + "view" ||
                    e.Name == ns + "zoom")
                .InDocumentOrder()
                .LastOrDefault();


            if (lastOfTop3 == null)
            {
                // none of those three exist, so add as first children of the root element
                settings.AddFirst(
                    settings.Elements(ns + "removePersonalInformation").Any() ?
                        null :
                        new XElement(ns + "removePersonalInformation"),
                    settings.Elements(ns + "removeDateAndTime").Any() ?
                        null :
                        new XElement(ns + "removeDateAndTime")
                );
            }
            else
            {
                // one of those three exist, so add after the last one
                lastOfTop3.AddAfterSelf(
                    settings.Elements(ns + "removePersonalInformation").Any() ?
                        null :
                        new XElement(ns + "removePersonalInformation"),
                    settings.Elements(ns + "removeDateAndTime").Any() ?
                        null :
                        new XElement(ns + "removeDateAndTime")
                );
            }
            using (XmlWriter xw =
              XmlWriter.Create(documentSettingsPart.GetStream(FileMode.Create, FileAccess.Write)))
                documentSettings.Save(xw);

            //Debug.WriteLine(documentSettings.ToString());

        }

        /// <summary>
        /// Removes personal informations from a Excel spreadsheet.
        /// </summary>
        /// <param name="document"></param>
        public static void RemovePersonalInfo(SpreadsheetDocument document)
        {
            RemovePersonalInfo(document.ExtendedFilePropertiesPart, document.CoreFilePropertiesPart);

            var workbookPart = document.WorkbookPart;
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            XNamespace x15ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";

            XDocument workbook = workbookPart.GetXDocument();

            //Debug.WriteLine(workbook.ToString());

            // remove absolute path
            workbook.Root.Elements(mc + "AlternateContent").Elements(mc + "Choice").Elements(x15ac + "absPath").Remove();
            // get workbookPr node to add filterPrivacy attribute
            XElement workbookPr = workbook.Root.Elements().Where(e => e.Name == ns + "workbookPr").FirstOrDefault();

            if (workbookPr == null)
            {
                workbookPr = new XElement(ns + "workbookPr");

                XElement fileVersion = workbook.Root.Elements().Where(e => e.Name == ns + "fileVersion").FirstOrDefault();
                if (fileVersion == null)
                {
                    workbook.Root.AddFirst(workbookPr);
                }
                else
                {
                    fileVersion.AddAfterSelf(workbookPr);
                }
            }
            workbookPr.SetAttributeValue("filterPrivacy", 1);

          using (XmlWriter xw =
              XmlWriter.Create(workbookPart.GetStream(FileMode.Create, FileAccess.Write)))
                workbook.Save(xw);

          //Debug.WriteLine(workbook.ToString());
        }

        /// <summary>
        /// Removes personal informations from a PowerPoint presentation.
        /// </summary>
        /// <param name="document"></param>
        public static void RemovePersonalInfo(PresentationDocument document)
        {
            RemovePersonalInfo(document.ExtendedFilePropertiesPart, document.CoreFilePropertiesPart);

            var presentationPart = document.PresentationPart;
            XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";

            XDocument presentation = presentationPart.GetXDocument();

            //Debug.WriteLine(presentation.ToString());

            presentation.Root.SetAttributeValue("removePersonalInfoOnSave", 1);

            using (XmlWriter xw =
                XmlWriter.Create(presentationPart.GetStream(FileMode.Create, FileAccess.Write)))
                presentation.Save(xw);

            //Debug.WriteLine(presentation.ToString());
        }

        /// <summary>
        /// Removes common personal informations from a Word document, Excel spreadsheet or PowerPoint presentation.
        /// </summary>
        /// <param name="extendedFilePropertiesPart"></param>
        /// <param name="coreFilePropertiesPart"></param>
        static void RemovePersonalInfo(
            OpenXmlPart extendedFilePropertiesPart,
            OpenXmlPart coreFilePropertiesPart
            )
        {
            // remove the company name from /docProps/app.xml
            // set TotalTime to "0"
            XNamespace x = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            XDocument extendedFilePropertiesXDoc = extendedFilePropertiesPart.GetXDocument();

            //Console.WriteLine(extendedFilePropertiesXDoc.ToString());

            extendedFilePropertiesXDoc.Elements(x + "Properties").Elements(x + "Company").Remove();
            XElement totalTime = extendedFilePropertiesXDoc.Elements(x + "Properties").Elements(x + "TotalTime").FirstOrDefault();
            if (totalTime != null)
                totalTime.Value = "0";
            using (XmlWriter xw =
              XmlWriter.Create(extendedFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                extendedFilePropertiesXDoc.Save(xw);


            // remove the values of dc:creator, cp:lastModifiedBy from /docProps/core.xml
            // set cp:revision to "1"
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            XDocument coreFilePropertiesXDoc = coreFilePropertiesPart.GetXDocument();

            //Debug.WriteLine(coreFilePropertiesXDoc.ToString());

            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + "coreProperties")
                                                           .Elements(dc + "creator")
                                                           .Nodes()
                                                           .OfType<XText>())
                textNode.Value = "";
            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + "coreProperties")
                                                           .Elements(cp + "lastModifiedBy")
                                                           .Nodes()
                                                           .OfType<XText>())
                textNode.Value = "";
            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + "coreProperties")
                                                       .Elements(dc + "title")
                                                       .Nodes()
                                                       .OfType<XText>())
                textNode.Value = "";
            XElement revision = coreFilePropertiesXDoc.Elements(cp + "coreProperties").Elements(cp + "revision").FirstOrDefault();
            if (revision != null)
                revision.Value = "1";
            using (XmlWriter xw =
              XmlWriter.Create(coreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                coreFilePropertiesXDoc.Save(xw);


        }

    }
}
