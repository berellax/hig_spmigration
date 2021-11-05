using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Text;

namespace HIGKnowledgePortal
{
    public class KnowledgePortal
    {
        public string HomeOffice { get; set; }
        public string DealStatus { get; set; }
        public string Name { get; set; }
        public string Created { get; set; }
        public string Country { get; set; }
        public string Industry { get; set; }
        public string Keywords { get; set; }
        public string Modified { get; set; }
        public string Fund { get; set; }
        public string DriveUrl { get; set; }
        public string CompanyName { get; set; }
        public string DriveId { get; set; }
        public string Key { get; set; }
        public string DocumentType { get; set; }
        public string DealTeam { get; set; }
        public string CreatedBy { get; set; }

        public KnowledgePortal()
        {

        }
        public KnowledgePortal (string[] row)
        {
            this.HomeOffice = row[0];
            this.DealStatus = row[1];
            this.Name = row[2];
            this.Created = row[3];
            this.Country = row[4];
            this.Industry = row[5];
            this.Keywords = row[6];
            this.Modified = row[7];
            this.Fund = row[8];
            this.DriveUrl = row[9];
            this.CompanyName = row[10];
            this.Key = row[11];
            this.DriveId = row[12];
            this.DocumentType = row[13];
            this.DealTeam = row[14];
        }

        public static List<KnowledgePortal> GetData()
        {
            List<KnowledgePortal> knowledgePortals = new List<KnowledgePortal>();
            using (TextFieldParser parser = new TextFieldParser(Configuration.MappingFilePath))
            {
                bool firstLine = true;
                //string[] headers;
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();

                    if (firstLine)
                    {
                        //headers = row;
                        firstLine = false;
                        continue;
                    }

                    knowledgePortals.Add(new KnowledgePortal(row));
                }
            }

            return knowledgePortals;
        }
    }
}
