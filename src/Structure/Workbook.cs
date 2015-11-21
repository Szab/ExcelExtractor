///-------------------------------------------------------------------------///
///   Namespace:      Szab.ExcelExtractor                                   ///
///   Class:          Workbook                                              ///
///   Description:    Representation of a whole excel workbook              ///
///   Author:         Szab                              Date: 20.11.2015    ///
///                                                                         ///
///   Notes:                                                                ///
///                                                                         ///
///                                                                         ///
///                                                                         ///
///-------------------------------------------------------------------------///

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Szab.Excel
{
    public class Workbook
    {
        private List<Sheet> _sheets;
        private List<string> _sharedStrings = new List<string>();

        public string Author
        {
            get;
            private set;
        }
        public string ModifiedBy
        {
            get;
            private set;
        }
        public DateTime? CreatedOn
        {
            get;
            private set;
        }
        public DateTime? LastModified
        {
            get;
            private set;
        }

        public Sheet this[int index]
        {
            get
            {
                return this.GetSheet(index);
            }
        }

        public Sheet[] Sheets 
        { 
            get
            {
                return _sheets.ToArray<Sheet>();
            }
        }

        public Workbook(string author = null, string createdOnISO = null, string modifiedBy = null, string modifiedISO = null, IEnumerable<string> sharedStrings = null)
        {
            this.Author = author;
            this.ModifiedBy = modifiedBy;

            if (sharedStrings != null)
            {
                _sharedStrings = sharedStrings.ToList<string>();
            }

            if (!string.IsNullOrEmpty(createdOnISO))
            {
                this.CreatedOn = DateTime.Parse(createdOnISO);
            }

            if (!string.IsNullOrEmpty(modifiedISO))
            {
                this.LastModified = DateTime.Parse(modifiedISO);
            }

            _sheets = new List<Sheet>();
        }

        public void AddSheet(Sheet sheet)
        {
            if(sheet == null)
            {
                throw new ArgumentNullException("You cannot add null sheet to the workbook");
            }

            this._sheets.Add(sheet);
        }

        public void AddSharedString(string sharedString)
        {
            this._sharedStrings.Add(sharedString);
        }

                public Sheet GetSheet(string sheetName)
        {
            return this._sheets.First(x => string.Equals(x.SheetName, sheetName));
        }


        public Sheet GetSheet(int index)
        {
            if(index >= 0 && index < this._sheets.Count)
            {
                return this._sheets[index];
            }   
            else
            {
                return null;
            }
        }

        public string GetSharedString(int index)
        {
            if(index >= 0 && index < this._sharedStrings.Count)
            {
                return this._sharedStrings[index];
            }
            else
            {
                return null;
            }
        }

        public void RemoveSheet(Sheet sheet)
        {
            this._sheets.Remove(sheet);
        }

        public void RemoveSheet(int index)
        {
            this._sheets.RemoveAt(index);
        }

        public void RemoveSharedString(string sharedString)
        {
            this._sharedStrings.Remove(sharedString);
        }

        public void RemoveSharedString(int index)
        {
            this._sharedStrings.RemoveAt(index);
        }

    }
}
