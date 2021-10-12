using System;
using PdfToExcel_covid19_yamagata.Dto;
using UglyToad.PdfPig;

namespace PdfToExcel_covid19_yamagata.IBiz
{
    public interface IPdfControler : IDisposable
    {
        PdfDocument PdfDoc { get; protected set; }

        void Load(string loadFilePath);
        PdfTextDataDto GetPdfTextData();
    }
}
