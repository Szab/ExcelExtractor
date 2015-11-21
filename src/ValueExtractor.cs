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

namespace Szab.Excel
{
    public class ValueExtractor
    {
        #region Private fields
        private Dictionary<string, string> _relationships;
        #endregion

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

        private void ParseWorkbook(string coreXml, string workbookXml, string sharedStringsXml, string relationshipsXml)
        {
            // Load document metadata
            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.LoadXml(coreXml);
            string author = this.FindNodeByNameRecursively("dc:creator", xmlDoc).InnerText;
            string modifiedBy = this.FindNodeByNameRecursively("cp:lastModifiedBy", xmlDoc).InnerText;
            string createdOn = this.FindNodeByNameRecursively("dcterms:created", xmlDoc).InnerText;
            string modifiedOn = this.FindNodeByNameRecursively("dcterms:modified", xmlDoc).InnerText;

            // Load relationships
            xmlDoc.LoadXml(relationshipsXml);
            XmlNode relationships = this.FindNodeByNameRecursively("Relationships", xmlDoc);

            foreach(XmlNode relationship in relationships)
            {
                string relationshipId = relationship.Attributes.GetNamedItem("Id").InnerText;
                string relationshipTarget = relationship.Attributes.GetNamedItem("Target").InnerText;

                this._relationships.Add(relationshipId, relationshipTarget);
            }

            // Load shared strings
            xmlDoc.LoadXml(sharedStringsXml);
            XmlNode sstNode = this.FindNodeByNameRecursively("sst", xmlDoc);
            List<string> sharedStrings = new List<string>();

            foreach (XmlNode sharedString in sstNode)
            {
                XmlNode textNode = this.FindNodeByNameRecursively("t", sharedString);

                if (textNode != null)
                {
                    sharedStrings.Add(textNode.InnerText);
                }
            }

            this.Workbook = new Workbook(author, createdOn, modifiedBy, modifiedOn, sharedStrings);

            // Load used sheets
            xmlDoc.LoadXml(workbookXml);
            XmlNode sheets = this.FindNodeByNameRecursively("sheets", xmlDoc);

            foreach(XmlNode sheet in sheets.ChildNodes)
            {
                string sheetName = sheet.Attributes.GetNamedItem("name").InnerText;
                string sheetId = sheet.Attributes.GetNamedItem("r:id").InnerText;

                this.Workbook.AddSheet(new Sheet(sheetName, sheetId));
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
                    XmlNode type = cell.Attributes.GetNamedItem("t");
                    string typeName = type != null ? type.InnerText : String.Empty;
                    string cellCoords = cell.Attributes.GetNamedItem("r").InnerText;
                    XmlNode cellValue = this.FindNodeByNameRecursively("v", cell);
                    string value = cellValue != null ? cellValue.InnerText : null;

                    if (!string.Equals(typeName, "s"))
                    {
                        sheet[cellCoords] = value;
                    }
                    else
                    {
                        int index = int.Parse(value);
                        sheet[cellCoords] = this.Workbook.GetSharedString(index);
                    }
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

        private string GetXmlFromFile(ZipArchiveEntry entry)
        {
            string result;

            using (Stream stream = entry.Open())
            {
                StreamReader streamReader = new StreamReader(stream);
                result = streamReader.ReadToEnd();
            }

            return result;
        }

        #endregion

        #region Public methods

        public ValueExtractor(string excelFilePath)
        {
            if(string.IsNullOrEmpty(excelFilePath))
            {
                throw new ArgumentNullException("Passing null paths as a constructor argument is not permitted.");
            }

            this._relationships = new Dictionary<string, string>();
            this.FilePath = excelFilePath;

            using(ZipArchive _excelFile = ZipFile.Open(excelFilePath, ZipArchiveMode.Read))
            {
                // Get metadata entries
                ZipArchiveEntry coreEntryFile = _excelFile.GetEntry("docProps/core.xml");
                ZipArchiveEntry workbookFile = _excelFile.Entries.First(x => string.Equals(x.Name, "workbook.xml"));
                ZipArchiveEntry sharedStringsFile = _excelFile.Entries.First(x => string.Equals(x.Name, "sharedStrings.xml"));
                ZipArchiveEntry relationshipsFile = _excelFile.GetEntry("xl/_rels/workbook.xml.rels");

                // Read metadata
                string workbookXml = this.GetXmlFromFile(workbookFile);
                string coreXml = this.GetXmlFromFile(coreEntryFile);
                string sharedStringsXml = this.GetXmlFromFile(sharedStringsFile);
                string relationshipsXml = this.GetXmlFromFile(relationshipsFile);


                this.ParseWorkbook(coreXml, workbookXml, sharedStringsXml, relationshipsXml);

                // Get all sheet files and populate existing sheet objects
                Sheet[] usedSheets = this.Workbook.Sheets;
                IEnumerable<ZipArchiveEntry> sheetFiles = _excelFile.Entries.Where(x => x.FullName.StartsWith("xl/worksheets"));
                
                foreach(Sheet sheet in usedSheets)
                {
                    ZipArchiveEntry sheetFile = sheetFiles.First(x => string.Equals(x.FullName, "xl/"+this._relationships[sheet.SheetId]));

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
