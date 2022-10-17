using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using TestExcelConsoleApp;



const string inputFileName = "Input\\ResourcesReport.xlsx";

const string inputTestFilePath = "Input\\InputTest.xlsx";
const string outputFileName = "Output\\OutputTest.xlsx";


string inputPath = Path.Combine(Directory.GetCurrentDirectory(), inputTestFilePath);
string outputPath = Path.Combine(Directory.GetCurrentDirectory(), outputFileName);


ExcelFilesReader reader = new ExcelFilesReader();
var inputBook = reader.ReadWorkbookXlsx(inputPath);

Console.WriteLine("Book: " + inputTestFilePath + '\n');
foreach (ISheet sheet in inputBook)
{
    Console.WriteLine("Sheet: " + sheet.SheetName);
    Console.WriteLine("[0][0]: " + sheet.GetRow(0).GetCell(0));
   
}

ExcelFilesWriter writer = new ExcelFilesWriter();


writer.ModifyCellFunctions(inputBook, outputPath);











