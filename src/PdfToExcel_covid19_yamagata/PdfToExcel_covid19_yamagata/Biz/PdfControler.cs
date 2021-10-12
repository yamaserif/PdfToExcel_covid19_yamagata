using System;
using PdfToExcel_covid19_yamagata.IBiz;
using PdfToExcel_covid19_yamagata.Dto;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfToExcel_covid19_yamagata.Biz
{
    public class PdfControler : IPdfControler
    {
        private string loadFilePath = null;
        private bool disposedValue = false;

        private PdfDocument pdfDoc = null;
        PdfDocument IPdfControler.PdfDoc
        { 
            get { return this.pdfDoc; }
            set { this.pdfDoc = value; }
        }

        public void Load(string loadFilePath)
        {
            this.loadFilePath = loadFilePath;
            this.pdfDoc = PdfDocument.Open(loadFilePath);
        }

        public PdfTextDataDto GetPdfTextData()
        {
            var pdfData = new PdfTextDataDto(this.loadFilePath);

            foreach (Page page in this.pdfDoc.GetPages())
            {
                pdfData.Texts.Add(page.Text);
            }

            return pdfData;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    this.pdfDoc?.Dispose();
                    this.pdfDoc = null;
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
