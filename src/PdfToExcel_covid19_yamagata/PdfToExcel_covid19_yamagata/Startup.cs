using System.Collections.Generic;
using System.Configuration;
using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.Biz;
using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata
{
    class Startup
    {
        static void Main(string[] args)
        {
            // This soft use libraries.
            // https://github.com/UglyToad/PdfPig
            // https://github.com/OfficeDev/Open-XML-SDK

            // pdfからテキストを抽出
            var pdfTexts = new List<PdfTextDataDto>();
            using (IPdfControler pdfControler = new PdfControler())
            {
                foreach (var item in args)
                {
                    pdfControler.Load(item);
                    pdfTexts.Add(pdfControler.GetPdfTextData());
                }
            }

            // テキストから感染者のデータを抽出
            var covid19Data = new List<Covid19DataDto>();
            IToCovid19DataParser toCovid19DataParser = new ToCovid19DataParser();
            foreach (var text in pdfTexts)
            {
                toCovid19DataParser.SetData(text);
                covid19Data.Add(toCovid19DataParser.GetCovid19Data());
            }

            // 感染者のデータをexcelに書き込み
            using (IExcelControler excelControler = new ExcelControler())
            {
                excelControler.Load(ConfigurationManager.AppSettings["ExcelFilePath"],
                                    ConfigurationManager.AppSettings["TargetSheetName"]);
                excelControler.WriteCovid19Data(covid19Data);
            }
        }
    }
}
