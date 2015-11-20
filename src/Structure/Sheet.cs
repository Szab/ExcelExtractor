///-------------------------------------------------------------------------///
///   Namespace:      Szab.ExcelExtractor                                   ///
///   Class:          Sheet                                                 ///
///   Description:    Representation of a single calculation sheet          ///
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
    public class Sheet
    {
        #region Private fields

        private Dictionary<string, string> _values;

        #endregion

        #region Properties

        public string SheetName
        {
            get;
            private set;
        }

        public int SheetId
        {
            get;
            private set;
        }

        public string this[string index]
        {
            get
            {
                return this.GetValue(index);
            }
            set
            {
                this.SetValue(index, value);
            }
        }

        public string this[char index1, int index2]
        {
            get
            {
                return this.GetValue(index1, index2);
            }

            set
            {
                this.SetValue(index1, index2, value);
            }
        }

        #endregion

        #region Methods

        public Sheet(string sheetName = null, int sheetId = -1)
        {
            this.SheetName = sheetName;
            this.SheetId = sheetId;
            this._values = new Dictionary<string, string>();
        }

        public string GetValue(string index)
        {
            if (_values.ContainsKey(index))
            {
                return this._values[index];
            }
            else
            {
                return null;
            }
        }

        public string GetValue(char index1, int index2)
        {
            string fullIndex = index1 + index2.ToString();

            return this.GetValue(fullIndex);
        }

        public void SetValue(string index, string value)
        {
            this._values[index] = value;
        }

        public void SetValue(char index1, int index2, string value)
        {
            string fullIndex = index1 + index2.ToString();

            this.SetValue(fullIndex, value);
        }

        #endregion Methods
    }
}
