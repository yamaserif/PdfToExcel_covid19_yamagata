namespace PdfToExcel_covid19_yamagata.Dto
{
    public class Covid19RowDataDto
    {
        public int? Number { get; set; } = null;
        public int? SubNumber { get; set; } = null;
        public string Age { get; set; } = null;
        public string Sex { get; set; } = null;
        public string Address { get; set; } = null;
        public string Relation { get; set; } = null;

        public Covid19RowDataDto()
        {
            // 処理なし
        }

        public Covid19RowDataDto(int? number, int? subNumber, string age, string sex, string address, string relation)
        {
            this.Number = number;
            this.SubNumber = subNumber;
            this.Age = age;
            this.Sex = sex;
            this.Address = address;
            this.Relation = relation;
        }
    }
}
