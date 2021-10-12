using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata.IBiz
{
    public interface IToCovid19DataParser
    {
        PdfTextDataDto PdfTextData { get; protected set; }

        void SetData(PdfTextDataDto pdfTextData);
        Covid19DataDto GetCovid19Data();
    }
}
