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
using System.Threading.Tasks;

namespace Szab.ExcelExtractor
{
    public class ExcelExtractor
    {
        private ZipArchive _excelFile;

        private bool ValidateExcelFile(ZipArchive archive)
        {
            // TODO
            return true;
        }

        public ExcelExtractor(ZipArchive excelFile)
        {
            if(excelFile == null)
            {
                throw new ArgumentNullException("Passing null archive as a constructor argument is not permitted.");
            } 
            else if(!ValidateExcelFile(excelFile))
            {
                throw new ArgumentException("Provided archive is not a valid MS Excel file.", "excelFile");
            }

            _excelFile = excelFile;
        }

        public ExcelExtractor(string excelFilePath)
        {
            if(string.IsNullOrEmpty(excelFilePath))
            {
                throw new ArgumentNullException("Passing null paths as a constructor argument is not permitted.");
            }

            _excelFile = ZipFile.Open(excelFilePath, ZipArchiveMode.Read);
        }
    }
}
