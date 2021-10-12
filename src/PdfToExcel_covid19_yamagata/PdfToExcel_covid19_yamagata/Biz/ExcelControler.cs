using System;
using System.Linq;
using System.Collections.Generic;
using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.LocalizedBiz;
using PdfToExcel_covid19_yamagata.Dto;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PdfToExcel_covid19_yamagata.Biz
{
    public class ExcelControler : IExcelControler
    {
        private bool disposedValue = false;

        private SpreadsheetDocument excelDoc = null;
        private string targetSheetName = null;

        SpreadsheetDocument IExcelControler.ExcelDoc
        { 
            get { return this.excelDoc; }
            set { this.excelDoc = value; }
        }
        string IExcelControler.TargetSheetName
        {
            get { return this.targetSheetName; }
            set { this.targetSheetName = value; }
        }

        public void Load(string loadFilePath, string targetSheetName)
        {
            this.excelDoc = SpreadsheetDocument.Open(loadFilePath, true);
            this.targetSheetName = targetSheetName;
        }

        public void WriteCovid19Data(IEnumerable<Covid19DataDto> writeData)
        {
            writeData.OrderBy(wd => wd.Date);
            foreach (var item in writeData)
            {
                Console.WriteLine(item.Path + " のデータをExcelの書き込んでいます。");
                IExcelLocalizedControler excelLocalizedControler = null;
                switch (item.Creator)
                {
                    case CREATOR.YamagataPref:
                        excelLocalizedControler = new ExcelYamagataPrefControler();
                        break;
                    case CREATOR.YamagataCity:
                        excelLocalizedControler = new ExcelYamagataCityControler();
                        break;
                    case CREATOR.Other:
                    default:
                        break;
                }

                excelLocalizedControler?.SetData(item, this.targetSheetName);
                excelLocalizedControler?.WriteCovid19Row(this.excelDoc);
            }
        }

        // https://docs.microsoft.com/ja-jp/office/open-xml/how-to-insert-text-into-a-cell-in-a-spreadsheet?redirectedfrom=MSDN
        // 上記ページを参照
        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    this.excelDoc?.Close();
                    this.excelDoc = null;
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
