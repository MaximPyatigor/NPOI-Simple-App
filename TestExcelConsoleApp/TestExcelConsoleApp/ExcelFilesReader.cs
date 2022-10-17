using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelConsoleApp
{
    internal class ExcelFilesReader
    {
        public IEnumerable<FileInfo> readFilesInfo(string directoryPath)
        {
            DirectoryInfo inputDirectory = new DirectoryInfo(directoryPath);

            var xlsxFiles = inputDirectory.GetFiles("*.xlsx");
            var xlsFiles = inputDirectory.GetFiles("*.xls");

            var allExcelFiles = xlsxFiles.Concat(xlsFiles);

            return allExcelFiles;
        }

        public IWorkbook ReadWorkbookXls(string path)
        {
            IWorkbook book;

            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

            book = new HSSFWorkbook(fs);

            return book;
        }

        public IWorkbook ReadWorkbookXlsx(string path)
        {
            IWorkbook book;

            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

            book = new XSSFWorkbook(fs);

            return book;
        }

    }
}
