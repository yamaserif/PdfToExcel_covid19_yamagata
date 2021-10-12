using System;
using System.Collections.Generic;

namespace PdfToExcel_covid19_yamagata.Dto
{
    public enum CREATOR
    {
        YamagataPref, YamagataCity, Other
    }

    public class Covid19DataDto
    {
        public string Path { get; set; } = null;
        public DateTime? Date { get; set; } = null;
        public CREATOR Creator { get; set; } = CREATOR.Other;
        public List<Covid19RowDataDto> Covid19Data = null;

        public Covid19DataDto()
        {
            this.Covid19Data = new List<Covid19RowDataDto>();
        }

        public Covid19DataDto(string path)
        {
            this.Path = path;
            this.Covid19Data = new List<Covid19RowDataDto>();
        }
    }
}
