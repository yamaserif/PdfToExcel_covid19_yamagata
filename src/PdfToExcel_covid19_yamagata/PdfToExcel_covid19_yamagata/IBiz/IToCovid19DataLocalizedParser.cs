using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata.IBiz
{
    public interface IToCovid19DataLocalizedParser
    {
        string[] Words { get; protected set; }

        void SetData(string[] words);
        void ParseData(Covid19DataDto baseData);
    }
}
