using PdfToExcel_covid19_yamagata.Dto;
using DocumentFormat.OpenXml.Packaging;

namespace PdfToExcel_covid19_yamagata.IBiz
{
    public interface IExcelLocalizedControler
    {
        Covid19DataDto WriteData { get; protected set; }
        string TargetSheetName { get; protected set; }

        void SetData(Covid19DataDto writeData, string targetSheetName);
        void WriteCovid19Row(SpreadsheetDocument excelDoc);
    }
}
