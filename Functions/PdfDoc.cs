using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace HIGKnowledgePortal
{
    public class PdfDoc
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public string Creator { get; set; }
        public DateTime CreationDate { get; set; }
        public string Keywords { get; set; }
        public bool Error { get; set; }

        public PdfDoc()
        {

        }

        public PdfDoc(string filePath)
        {
            try
            {
                PdfDocument pdf = PdfReader.Open(filePath, PdfDocumentOpenMode.ReadOnly);

                if(pdf != null)
                {
                    this.Title = pdf.Info.Title;
                    this.Author = pdf.Info.Author;
                    this.Creator = pdf.Info.Creator;
                    this.CreationDate = pdf.Info.CreationDate;
                    this.Keywords = pdf.Info.Keywords;
                }
            }
            catch (Exception ex)
            {
                this.Error = true;
                Console.WriteLine("Error opening PDF Document {0}. Message {1}.", filePath, ex.Message);
            }
        }

        public static string[] GetPdfDocumentsFromDirectory(string directoryName)
        {
            string[] fileNames;

            if (!Directory.Exists(directoryName))
                throw new Exception($"Directory Name {directoryName} does not exist");

            fileNames = Directory.GetFiles(directoryName, "*.pdf");

            return fileNames;
        }

        public static string ExportPdfInfoToCsv(List<PdfDoc> pdfDocs)
        {
            string outputFile = $"{Configuration.DownloadDirectory}\\AuthorOutput.csv";

            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            using (StreamWriter writer = new StreamWriter(new FileStream(outputFile, FileMode.Create, FileAccess.Write)))
            {
                char delimeter = '|';
                string newLine = Environment.NewLine;

                string[] header = { "FileName", "CreatedDate", "Author", "Creator" };
                writer.WriteLine(String.Join(delimeter, header));

                foreach (PdfDoc pdf in pdfDocs)
                {
                    string[] line = { pdf.Title, pdf.CreationDate.ToString(), pdf.Author, pdf.Creator };
                    writer.WriteLine(String.Join(delimeter, line));

                    Console.WriteLine("--------------------------------------------------------------");
                    Console.WriteLine("Title: {0}", pdf.Title);
                    Console.WriteLine("CreationDate: {0}", pdf.CreationDate);
                    Console.WriteLine("Author: {0}", pdf.Author);
                    Console.WriteLine("Creator: {0}", pdf.Creator);
                }

                writer.Close();
            }

            return outputFile;
        }
    }
}
