using System;
using System.Linq;
using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.Biz;
using PdfToExcel_covid19_yamagata.Dto;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PdfToExcel_covid19_yamagata.LocalizedBiz
{
    public class ExcelYamagataPrefControler : IExcelLocalizedControler
    {
        private Covid19DataDto writeData = null;
        private string targetSheetName = null;

        Covid19DataDto IExcelLocalizedControler.WriteData
        {
            get { return this.writeData; }
            set { this.writeData = value; }
        }
        string IExcelLocalizedControler.TargetSheetName
        {
            get { return this.targetSheetName; }
            set { this.targetSheetName = value; }
        }

        public void SetData(Covid19DataDto writeData, string targetSheetName)
        {
            this.writeData = writeData;
            this.targetSheetName = targetSheetName;
        }

        public void WriteCovid19Row(SpreadsheetDocument excelDoc)
        {
            var wbPart = excelDoc.WorkbookPart;
            var targetSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == this.targetSheetName)
                                                                  .FirstOrDefault();

            if(targetSheet != null)
            {
                var wsPart = (WorksheetPart)wbPart.GetPartById(targetSheet.Id);
                var checkCell = wsPart.Worksheet.Descendants<Cell>().Where(s => s.CellReference?.Value?.StartsWith('Q') ?? false)
                                                                    .LastOrDefault();

                if (checkCell != null)
                {

                    var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                    var lastIndex = uint.Parse(checkCell.CellReference?.Value?.TrimStart('Q'));
                    var writeIndex = lastIndex + 2;

                    foreach (var writeItem in this.writeData.Covid19Data)
                    {
                        var cellQ = ExcelControler.InsertCellInWorksheet("Q", writeIndex, wsPart);
                        cellQ.CellValue = new CellValue(this.writeData.Date?.ToString("yyyy/MM/dd") ?? string.Empty);
                        cellQ.DataType = CellValues.String;

                        var cellS = ExcelControler.InsertCellInWorksheet("S", writeIndex, wsPart);
                        cellS.CellValue = new CellValue(writeItem.Number ?? 0);
                        cellS.DataType = CellValues.Number;

                        var cellT = ExcelControler.InsertCellInWorksheet("T", writeIndex, wsPart);
                        cellT.CellValue = new CellValue(writeItem.Age);
                        cellT.DataType = CellValues.String;

                        var cellU = ExcelControler.InsertCellInWorksheet("U", writeIndex, wsPart);
                        cellU.CellValue = new CellValue(writeItem.Sex);
                        cellU.DataType = CellValues.String;

                        var cellV = ExcelControler.InsertCellInWorksheet("V", writeIndex, wsPart);
                        cellV.CellValue = new CellValue(writeItem.Address);
                        cellV.DataType = CellValues.String;

                        var cellAB = ExcelControler.InsertCellInWorksheet("AB", writeIndex, wsPart);
                        cellAB.CellValue = new CellValue(writeItem.Relation);
                        cellAB.DataType = CellValues.String;

                        writeIndex++;
                    }
                }
            }
        }
    }
}
