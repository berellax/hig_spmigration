using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace KeebTalentBook
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Select the routine that you would like to run.");
            Console.WriteLine("(1) Update Created Date.");
            Console.WriteLine("(2) Update Author Value.");
            Console.WriteLine("(3) Export Author Information");

            ConsoleKey response = Console.ReadKey(false).Key;

            if(response == ConsoleKey.D1)
            {
                UpdateCreatedDate();
            }
            else if(response == ConsoleKey.D2)
            {
                UpdateAuthor();
            }
            else if(response == ConsoleKey.D3)
            {
                ExportAuthorInfo();
            }
            else
            {
                return;
            }
 

            Console.WriteLine("Application has completed. Press any key to exit");
            Console.ReadLine();
        }

        private static void UpdateCreatedDate()
        {
            List<KnowledgePortal> knowledgePortal = KnowledgePortal.GetData();
            Console.WriteLine($"Text file parsed with {knowledgePortal.Count} items.");

            GraphApi graphApi = new GraphApi();

            string siteId = graphApi.GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string listId = graphApi.GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            SPListDefinition listColumns = graphApi.GetListDefinition(siteId, listId);

            SPColumnDefinition createdOnColumn = graphApi.GetColumnDefinition(listColumns, "created");
            Console.WriteLine($"Created On Column Id: {createdOnColumn.id}");

            if (createdOnColumn.readOnly)
            {
                graphApi.SetColumnReadOnly(siteId, listId, createdOnColumn.id, "false");
                Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag has been set to false");
            }
            else
            {
                Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag is already false");
            }

            List<SPListItem> listItems = graphApi.GetListItems(siteId, listId);
            Console.WriteLine($"List {Configuration.ListName} has {listItems.Count} items with a Title.");

            foreach (var item in listItems)
            {
                var matchItem = knowledgePortal.Where(a => a.Name.ToLower() == item.fields.Title.ToLower()).FirstOrDefault();
                if (matchItem != null)
                {
                    var createdOn = matchItem.Created;
                    string createdBy = null;

                    if (Configuration.SetAuthor)
                    {
                        matchItem.CreatedBy = "backup@keeebbob.onmicrosoft.com";
                        createdBy = graphApi.GetCreatedByValue(matchItem, siteId);
                    }

                    graphApi.UpdateListItem(siteId, listId, item.id, createdOn, createdBy);
                    Console.WriteLine($"Updated Item with ID: {item.id} | Name: {item.fields.Title} | Created: {createdOn}");
                }
                else
                {
                    Console.WriteLine($"Skipped Item with ID: {item.id} | Name: {item.fields.Title} | Not found in CSV Data.");
                }
            }


            graphApi.SetColumnReadOnly(siteId, listId, createdOnColumn.id, "true");
            Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag has been set to true");
        }

        private static void ExportAuthorInfo()
        {
            GraphApi graphApi = new GraphApi();

            string siteId = graphApi.GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string listId = graphApi.GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            string driveId = graphApi.GetDriveId(siteId);
            Console.WriteLine($"Drive ID: {driveId}");

            var driveItems = graphApi.GetDriveItems(siteId, driveId);
            Console.WriteLine($"Retrieved Drive Items");

            graphApi.DownloadFiles(driveItems);
            GetPdfFileInfo();
        }

        private static void UpdateAuthor()
        {
            throw new NotImplementedException();

            GraphApi graphApi = new GraphApi();

            string siteId = graphApi.GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string driveId = graphApi.GetDriveId(siteId);
            Console.WriteLine($"Drive ID: {driveId}");

            string listId = graphApi.GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            var driveItems = graphApi.GetDriveItems(siteId, driveId);
            Console.WriteLine($"Retrieved Drive Items");

            graphApi.DownloadFiles(driveItems);
            GetPdfFileInfo();

            SPListDefinition listColumns = graphApi.GetListDefinition(siteId, listId);
            SPColumnDefinition createdByColumn = new SPColumnDefinition();

            createdByColumn = graphApi.GetColumnDefinition(listColumns, "author");
            Console.WriteLine($"Created By Column Id: {createdByColumn.id}");

            if (createdByColumn.readOnly)
            {
                graphApi.SetColumnReadOnly(siteId, listId, createdByColumn.id, "false");
                Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag has been set to false");
            }
            else
            {
                Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag is already false");
            }

            //foreach (var item in listItems)
            //{
            //    var matchItem = knowledgePortal.Where(a => a.Name.ToLower() == item.fields.Title.ToLower()).FirstOrDefault();
            //    if (matchItem != null)
            //    {
            //        string createdBy = null;

            //        if (Configuration.SetAuthor)
            //        {
            //            matchItem.CreatedBy = "backup@keeebbob.onmicrosoft.com";
            //            createdBy = graphApi.GetCreatedByValue(matchItem, siteId);
            //        }

            //        graphApi.UpdateListItem(siteId, listId, item.id, createdOn, createdBy);
            //        Console.WriteLine($"Updated Item with ID: {item.id} | Name: {item.fields.Title} | Created: {createdOn}");
            //    }
            //    else
            //    {
            //        Console.WriteLine($"Skipped Item with ID: {item.id} | Name: {item.fields.Title} | Not found in CSV Data.");
            //    }
            //}

            graphApi.SetColumnReadOnly(siteId, listId, createdByColumn.id, "true");
            Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag has been set to true");
        }



        private static void GetPdfFileInfo()
        {
            if (!Directory.Exists(Configuration.DownloadDirectory))
                return;

            string[] pdfFileNames = Directory.GetFiles(Configuration.DownloadDirectory, "*.pdf");

            string filePath = $"{Configuration.DownloadDirectory}\\AuthorOutput.csv";

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (StreamWriter writer = new StreamWriter(new FileStream(filePath, FileMode.Create, FileAccess.Write)))
            {
                char delimeter = ',';
                string newLine = Environment.NewLine;

                string[] header = { "FileName", "CreatedDate", "Author", "Creator" };
                writer.WriteLine(String.Join(delimeter, header));

                foreach (string fileName in pdfFileNames)
                {
                    FileInfo file = new FileInfo(fileName);

                    if(file.Length > 0)
                    {
                        try
                        {
                            PdfDocument pdf = PdfReader.Open(fileName, PdfDocumentOpenMode.ReadOnly);

                            Console.WriteLine("--------------------------------------------------------------");
                            Console.WriteLine("Author: {0}", pdf.Info.Author);
                            Console.WriteLine("CreationDate: {0}", pdf.Info.CreationDate);
                            Console.WriteLine("Creator: {0}", pdf.Info.Creator);
                            Console.WriteLine("Keywords: {0}", pdf.Info.Keywords);

                            string[] line = { pdf.Info.Title, pdf.Info.CreationDate.ToString(), pdf.Info.Author, pdf.Info.Creator };
                            writer.WriteLine(String.Join(delimeter, line));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error opening PDF Document {fileName}. {ex}");
                        }
                    }

                }

                writer.Close();
            }

        }
    }

    



    
}
