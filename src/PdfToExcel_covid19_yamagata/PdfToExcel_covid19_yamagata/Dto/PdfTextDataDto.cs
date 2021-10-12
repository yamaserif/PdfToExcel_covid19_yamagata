using System.Collections.Generic;

namespace PdfToExcel_covid19_yamagata.Dto
{
    public class PdfTextDataDto
    {
        public string Path { get; set; } = null;
        public IList<string> Texts { get; set; } = null;

        public PdfTextDataDto()
        {
            // 処理なし
        }

        public PdfTextDataDto(string path)
        {
            this.Path = path;
            this.Texts = new List<string>();
        }
    }
}
