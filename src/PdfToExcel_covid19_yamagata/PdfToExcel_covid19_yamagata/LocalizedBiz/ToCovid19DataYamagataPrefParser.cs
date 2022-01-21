using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.Dto;
using System.Configuration;
using System.Linq;

namespace PdfToExcel_covid19_yamagata.LocalizedBiz
{
    public class ToCovid19DataYamagataPrefParser : IToCovid19DataLocalizedParser
    {
        private string[] relationLinefeedWords = ConfigurationManager.AppSettings["YamagataPref_RelationLinefeedWords"].Split(",");

        private string[] words = null;
        string[] IToCovid19DataLocalizedParser.Words
        {
            get { return this.words; }
            set { this.words = value; }
        }

        public void SetData(string[] words)
        {
            this.words = words;
        }

        public void ParseData(Covid19DataDto baseData)
        {
            Covid19RowDataDto row = null;
            for (var index = 0; index < this.words.Length; index++)
            {
                if(int.TryParse(this.words[index], out var no))
                {
                    if (index > 0 &&
                        (this.words[index - 1].EndsWith('/') ||
                         this.words[index + 1].StartsWith('/')))
                    {
                        continue;
                    }

                    if (row != null)
                    {
                        row.Relation = RelationWord(this.words[index - 1], this.words[index - 2]);
                        baseData.Covid19Data.Add(row);
                    }
                    row = new Covid19RowDataDto();
                    row.Number = no;
                    row.Age = this.words[++index];
                    row.Sex = this.words[++index];
                    row.Address = this.words[++index];
                }
                else if(this.words[index].StartsWith("〇"))
                {
                    if (row != null)
                    {
                        row.Relation = RelationWord(this.words[index - 1], this.words[index - 2]);
                        baseData.Covid19Data.Add(row);
                    }
                    break;
                }
                else if (index == this.words.Length - 1)
                {
                    if (row != null)
                    {
                        row.Relation = RelationWord(this.words[index], this.words[index - 1]);
                        baseData.Covid19Data.Add(row);
                    }
                    break;
                }
            }
        }

        private string RelationWord(string word, string beforeWord)
        {
            var returnWord = word;
            if (relationLinefeedWords.Contains(word))
            {
                returnWord = beforeWord + word;
            }
            return returnWord;
        }
    }
}
