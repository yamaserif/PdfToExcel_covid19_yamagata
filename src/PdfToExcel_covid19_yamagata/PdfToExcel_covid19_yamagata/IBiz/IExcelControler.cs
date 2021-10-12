using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata.IBiz
{
    public interface IExcelControler : IDisposable
    {
        SpreadsheetDocument ExcelDoc { get; protected set; }
        string TargetSheetName { get; protected set; }

        void Load(string loadFilePath, string targetSheetName);
        void WriteCovid19Data(IEnumerable<Covid19DataDto> writeData);
    }
}
