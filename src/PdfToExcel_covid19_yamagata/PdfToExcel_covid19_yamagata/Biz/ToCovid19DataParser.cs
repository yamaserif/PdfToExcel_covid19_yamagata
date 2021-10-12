using System;
using System.Globalization;
using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.LocalizedBiz;
using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata.Biz
{
    public class ToCovid19DataParser : IToCovid19DataParser
    {
        private PdfTextDataDto pdfTextData = null;
        PdfTextDataDto IToCovid19DataParser.PdfTextData
        {
            get { return this.pdfTextData; }
            set { this.pdfTextData = value; }
        }

        public void SetData(PdfTextDataDto pdfTextData)
        {
            this.pdfTextData = pdfTextData;
        }

        public Covid19DataDto GetCovid19Data()
        {
            var returnCovid19Data = new Covid19DataDto(this.pdfTextData.Path);

            foreach (var text in this.pdfTextData.Texts)
            {
                this.ParseData(text, returnCovid19Data);
            }

            return returnCovid19Data;
        }

        private void ParseData(string parseText, Covid19DataDto baseData)
        {
            var words = parseText.Trim()
                                 .Split(' ', StringSplitOptions.RemoveEmptyEntries);

            if(baseData.Date == null)
            {
                baseData.Date = this.GetDateTimeForDateText(words[0], baseData.Path);
                baseData.Creator = this.GetCreator(words[1], baseData.Path);
            }

            IToCovid19DataLocalizedParser toCovid19DataLocalizedParser = null;
            switch (baseData.Creator)
            {
                case CREATOR.YamagataPref:
                    toCovid19DataLocalizedParser = new ToCovid19DataYamagataPrefParser();
                    break;
                case CREATOR.YamagataCity:
                    toCovid19DataLocalizedParser = new ToCovid19DataYamagataCityParser();
                    break;
                case CREATOR.Other:
                default:
                    Console.WriteLine("PDFファイルの生成団体が確認できません。");
                    break;
            }

            toCovid19DataLocalizedParser?.SetData(words);
            toCovid19DataLocalizedParser?.ParseData(baseData);
        }

        private DateTime? GetDateTimeForDateText(string dateText, string filePath = "確認中のPDF")
        {
            DateTime? returnDateTime = null;

            var culture = new CultureInfo("ja-JP", true);
            culture.DateTimeFormat.Calendar = new JapaneseCalendar();

            if (DateTime.TryParseExact(this.ReplaceNumZenToHan(dateText), "ggyy年M月d日", culture,
                                       DateTimeStyles.AssumeLocal, out var tryParseExactResult))
            {
                returnDateTime = tryParseExactResult;
            }
            else
            {
                Console.WriteLine(filePath + " の日付が確認できません。");
                Console.WriteLine("「yyyy/MM/dd」の形で入力してください(Datetimeが認識できる形)");

                if (DateTime.TryParse(Console.ReadLine(), out var tryParseResult))
                {
                    returnDateTime = tryParseResult;
                }
                else
                {
                    Console.WriteLine("日付の設定に失敗しました。");
                }
            }

            return returnDateTime;
        }

        private string ReplaceNumZenToHan(string replaceText)
        {
            return replaceText.Replace('０', '0')
                              .Replace('１', '1')
                              .Replace('２', '2')
                              .Replace('３', '3')
                              .Replace('４', '4')
                              .Replace('５', '5')
                              .Replace('６', '6')
                              .Replace('７', '7')
                              .Replace('８', '8')
                              .Replace('９', '9');
        }

        private CREATOR GetCreator(string CreatorText, string filePath = "確認中のPDF")
        {
            var returnCreator = CREATOR.Other;

            if (CreatorText.Contains("山形県"))
            {
                returnCreator = CREATOR.YamagataPref;
            }
            else if (CreatorText.Contains("山形市"))
            {
                returnCreator = CREATOR.YamagataCity;
            }
            else
            {
                Console.WriteLine(filePath + " のPDFファイルの生成団体が確認できません。");
                Console.WriteLine("「山形県:0, 山形市:1」の形で入力してください");
                var inputCreatorText = Console.ReadLine();

                if (inputCreatorText == "0")
                {
                    returnCreator = CREATOR.YamagataPref;
                }
                else if (inputCreatorText == "1")
                {
                    returnCreator = CREATOR.YamagataCity;
                }
            }

            return returnCreator;
        }
    }
}
