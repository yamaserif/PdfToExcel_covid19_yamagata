using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.Dto;

namespace PdfToExcel_covid19_yamagata.LocalizedBiz
{
    public class ToCovid19DataYamagataCityParser : IToCovid19DataLocalizedParser
    {
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
                if(int.TryParse(this.words[index], out var subNo))
                {
                    if (index > 0 &&
                        (this.words[index - 1].EndsWith('/') ||
                         this.words[index + 1].StartsWith('/')))
                    {
                        continue;
                    }

                    if (row != null)
                    {
                        row.Relation = this.words[index - 1];
                        baseData.Covid19Data.Add(row);
                    }
                    row = new Covid19RowDataDto();
                    row.SubNumber = subNo;
                    if (int.TryParse(this.words[++index].Trim('(', ')'), out var no))
                    {
                        row.Number = no;
                    }
                    row.Age = this.words[++index];
                    row.Sex = this.words[++index];
                    row.Address = this.words[++index];
                }
                else if(this.words[index].StartsWith("○"))
                {
                    if (row != null)
                    {
                        row.Relation = this.words[index - 1];
                        baseData.Covid19Data.Add(row);
                    }
                    break;
                }
            }
        }
    }
}
