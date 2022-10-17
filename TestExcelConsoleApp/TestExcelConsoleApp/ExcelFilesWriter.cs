using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
//using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelConsoleApp
{
    internal class ExcelFilesWriter
    {
        public void WriteWorkbookToFile(IWorkbook sourceBook, string path)
        {
            using (var file = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                sourceBook.Write(file);
                file.Close();
            }
        }

        //public void WriteCellsTest(IWorkbook sourceBook, IWorkbook resultBook)
        //{

        //    for(int i = 0; i < sourceBook.NumberOfSheets; i++)
        //    {
        //        for(int j = 0; j < sourceBook.GetSheetAt(i).PhysicalNumberOfRows; j++)
        //        {
        //            //for Console.WriteLine(sourceBook.GetSheetAt(i).GetRow());
        //        }
        //    }
        //        //sourceBook.GetSheetAt(0).GetRow(0).GetCell(0).SetCellValue(sourceBook.);
        //}

        public IWorkbook ModifyCellFunctions(IWorkbook hssfwb, string path)
        {

            ISheet sheet = hssfwb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);


            ICellStyle testeStyle = hssfwb.CreateCellStyle();
            testeStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            testeStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;
            testeStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
            testeStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;

            testeStyle.TopBorderColor = IndexedColors.BrightGreen.Index;
            testeStyle.BottomBorderColor = IndexedColors.BrightGreen.Index;
            testeStyle.LeftBorderColor = IndexedColors.BrightGreen.Index;
            testeStyle.RightBorderColor = IndexedColors.BrightGreen.Index;

            ICell cell = row.CreateCell(5);
            cell.SetCellValue("testeeerere");
            cell.CellStyle = testeStyle;


            IDrawing patr = sheet.CreateDrawingPatriarch();

            //anchor defines size and position of the comment in worksheet
            IComment comment1 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));

            // set text in the comment
            comment1.String = new XSSFRichTextString("We can set comments in POI");

            //set comment author.
            //you can see it in the status bar when moving mouse over the commented cell
            comment1.Author = "Apache Software Foundation";


            cell.CellComment = comment1;    


            using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                hssfwb.Write(file);
                file.Close();
            }

            return null;

        }
    }
}
