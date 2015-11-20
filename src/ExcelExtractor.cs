///-------------------------------------------------------------------------///
///   Namespace:      Szab.ExcelExtractor                                   ///
///   Class:          ExcelExtractor                                        ///
///   Description:    A lib for extracting values from MSExcel files        ///
///   Author:         Szab                              Date: 20.11.2015    ///
///                                                                         ///
///   Notes:                                                                ///
///                                                                         ///
///                                                                         ///
///                                                                         ///
///-------------------------------------------------------------------------///

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace Szab.ExcelExtractor
{
    public class ExcelExtractor
    {
        public readonly string FilePath;
        public Workbook Workbook
        {
            get;
            private set;
        }

        private bool ValidateExcelFile(ZipArchive archive)
        {
            // TODO
            return true;
        }

        private void ParseWorkbook(string coreXml, string workbookXml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(coreXml);

            string author = this.FindNodeByNameRecursively("dc:creator", xmlDoc.LastChild).InnerText;
            string modifiedBy = this.FindNodeByNameRecursively("cp:lastModifiedBy", xmlDoc.LastChild).InnerText;
            string createdOn = this.FindNodeByNameRecursively("dcterms:created", xmlDoc.LastChild).InnerText;
            string modifiedOn = this.FindNodeByNameRecursively("dcterms:modified", xmlDoc.LastChild).InnerText;

            this.Workbook = new Workbook(author, createdOn, modifiedBy, modifiedOn);
        }

        private XmlNode FindNodeByNameRecursively(string name, XmlNode node)
        {
            XmlNode result = null;

            if(string.Equals(node.Name, name))
            {
                result = node;
            }
            else
            {
                foreach(XmlNode child in node)
                {
                    result = FindNodeByNameRecursively(name, child);

                    if (result != null)
                        break;
                }
            }

            return result;
        }

        public ExcelExtractor(string excelFilePath)
        {
            if(string.IsNullOrEmpty(excelFilePath))
            {
                throw new ArgumentNullException("Passing null paths as a constructor argument is not permitted.");
            }

            this.FilePath = excelFilePath;

            using(ZipArchive _excelFile = ZipFile.Open(excelFilePath, ZipArchiveMode.Read))
            {
                ZipArchiveEntry coreEntryFile = _excelFile.GetEntry("docProps/core.xml");
                ZipArchiveEntry workbookFile = _excelFile.Entries.First(x => string.Equals(x.Name, "workbook.xml"));

                string workbookXml;
                string coreXml;

                using (Stream workbookStream = workbookFile.Open())
                {
                    StreamReader workbookStreamReader = new StreamReader(workbookStream);
                    workbookXml = workbookStreamReader.ReadToEnd();
                }

                using (Stream coreStream = coreEntryFile.Open())
                {
                    StreamReader coreXmlStreamReader = new StreamReader(coreStream);
                    coreXml = coreXmlStreamReader.ReadToEnd();
                }

                this.ParseWorkbook(coreXml, workbookXml);
                
            }
        }
    }
}
