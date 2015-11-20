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
    class Workbook
    {
        private List<Sheet> _sheets;

        public readonly string Author;
        public readonly string ModifiedBy;
        public readonly DateTime? CreatedOn;
        public readonly DateTime? LastModified;
        public readonly string OriginalFilePath;

        public List<Sheet> Sheets 
        { 
            get
            {
                return _sheets;
            }
        }

        public Workbook(string path = null, string author = null, string createdOnISO = null, string modifiedBy = null, string modifiedISO = null)
        {
            this.Author = author;
            this.ModifiedBy = modifiedBy;

            if(!string.IsNullOrEmpty(createdOnISO))
            {
                this.CreatedOn = DateTime.Parse(createdOnISO);
            }

            if (!string.IsNullOrEmpty(modifiedISO))
            {
                this.CreatedOn = DateTime.Parse(modifiedISO);
            }

            OriginalFilePath = null;
            _sheets = new List<Sheet>();
        }
    }
}
