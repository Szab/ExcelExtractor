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

namespace Szab.ExcelExtractor
{
    public class Workbook
    {
        private List<Sheet> _sheets;

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

        public Workbook(string author = null, string createdOnISO = null, string modifiedBy = null, string modifiedISO = null)
        {
            this.Author = author;
            this.ModifiedBy = modifiedBy;

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

        public Sheet GetSheet(string sheetName)
        {
            return this._sheets.First(x => string.Equals(x.SheetName, sheetName));
        }

        public void AddSheet(Sheet sheet)
        {
            if(sheet == null)
            {
                throw new ArgumentNullException("You cannot add null sheet to the workbook");
            }

            this._sheets.Add(sheet);
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

        public void RemoveSheet(Sheet sheet)
        {
            this._sheets.Remove(sheet);
        }

        public void RemoveSheet(int index)
        {
            this._sheets.RemoveAt(index);
        }


    }
}
