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
        #region Properties

        public readonly string FilePath;
        public Workbook Workbook
        {
            get;
            private set;
        }

        #endregion

        #region Private methods

        private bool ValidateExcelFile(ZipArchive archive)
        {
            // TODO
            return true;
        }

        private void ParseWorkbook(string coreXml, string workbookXml)
        {
            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.LoadXml(coreXml);
            string author = this.FindNodeByNameRecursively("dc:creator", xmlDoc).InnerText;
            string modifiedBy = this.FindNodeByNameRecursively("cp:lastModifiedBy", xmlDoc).InnerText;
            string createdOn = this.FindNodeByNameRecursively("dcterms:created", xmlDoc).InnerText;
            string modifiedOn = this.FindNodeByNameRecursively("dcterms:modified", xmlDoc).InnerText;

            this.Workbook = new Workbook(author, createdOn, modifiedBy, modifiedOn);

            xmlDoc.LoadXml(workbookXml);
            XmlNode sheets = this.FindNodeByNameRecursively("sheets", xmlDoc);

            foreach(XmlNode sheet in sheets.ChildNodes)
            {
                string sheetName = sheet.Attributes.GetNamedItem("name").InnerText;
                string sheetId = sheet.Attributes.GetNamedItem("sheetId").InnerText;
                int sheetIdInt = -1;
                int.TryParse(sheetId, out sheetIdInt);

                this.Workbook.AddSheet(new Sheet(sheetName, sheetIdInt));
            }
        }

        private void PopulateSheet(Sheet sheet, string sheetXml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(sheetXml);

            XmlNode sheetData = this.FindNodeByNameRecursively("sheetData", xmlDoc);
            
            foreach(XmlNode row in sheetData)
            {
                foreach(XmlNode cell in row)
                {
                    string cellCoords = cell.Attributes.GetNamedItem("r").InnerText;
                    string value = this.FindNodeByNameRecursively("v", cell).InnerText;

                    sheet[cellCoords] = value;
                }
            }
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

        #endregion

        #region Public methods

        public ExcelExtractor(string excelFilePath)
        {
            if(string.IsNullOrEmpty(excelFilePath))
            {
                throw new ArgumentNullException("Passing null paths as a constructor argument is not permitted.");
            }

            this.FilePath = excelFilePath;

            using(ZipArchive _excelFile = ZipFile.Open(excelFilePath, ZipArchiveMode.Read))
            {
                // Get metadata entries
                ZipArchiveEntry coreEntryFile = _excelFile.GetEntry("docProps/core.xml");
                ZipArchiveEntry workbookFile = _excelFile.Entries.First(x => string.Equals(x.Name, "workbook.xml"));

                // Read metadata
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

                // Get all sheet files and populate existing sheet objects
                Sheet[] usedSheets = this.Workbook.Sheets;
                IEnumerable<ZipArchiveEntry> sheetFiles = _excelFile.Entries.Where(x => x.FullName.StartsWith("xl/worksheets"));
                
                foreach(Sheet sheet in usedSheets)
                {
                    ZipArchiveEntry sheetFile = sheetFiles.First(x => string.Equals(x.Name, "sheet"+sheet.SheetId+".xml"));

                    using (Stream sheetStream = sheetFile.Open())
                    {
                        StreamReader sheetXmlReader = new StreamReader(sheetStream);
                        this.PopulateSheet(sheet, sheetXmlReader.ReadToEnd());
                    }
                }
                
            }
        }

        #endregion
    }
}
