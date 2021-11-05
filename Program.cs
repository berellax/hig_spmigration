﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace KeebTalentBook
{
    public class Program
    {
        private static GraphApi _graphApi;

        static void Main(string[] args)
        {
            Console.WriteLine("Select the script that you would like to run.");
            Console.WriteLine("(1) Update Created Date");
            Console.WriteLine("(2) Update Author Value");
            Console.WriteLine("(3) Export Author Information");

            ConsoleKey response = Console.ReadKey(false).Key;

            //Instantiate Graph API
            AuthenticateGraphApi();

            //Execute method based on option selected
            if (response == ConsoleKey.D1)
            {
                UpdateCreatedDate();
            }
            else if(response == ConsoleKey.D2)
            {
                Console.WriteLine("This script is not available.");
                Environment.Exit(-1);
                //UpdateAuthor();
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

        #region Primary Methods
        private static void AuthenticateGraphApi()
        {
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Select the authentication method for Graph API.");
            Console.WriteLine("(1) App Registration (Client Id / Client Secret)");
            Console.WriteLine("(2) User Authentication (Username / Password");
            Console.WriteLine("(3) Authentication Token (Copy from Graph Explorer)");

            ConsoleKey response = Console.ReadKey(false).Key;
            Console.WriteLine();
            Console.WriteLine();

            switch (response)
            {
                case ConsoleKey.D1:
                    _graphApi = new GraphApi(AuthenticationType.AppRegistration);
                    break;
                case ConsoleKey.D2:
                    _graphApi = new GraphApi(AuthenticationType.UserAuthentication);
                    break;
                case ConsoleKey.D3:
                    _graphApi = new GraphApi(AuthenticationType.AuthenticationToken);
                    break;
                default:
                    Console.WriteLine("Invalid Entry. Please try again.");
                    AuthenticateGraphApi();
                    break;
            }

            Console.WriteLine("Graph API authentication complete.");

        }
        /// <summary>
        /// Updates the Created Date of items in the list based on the provided CSV file
        /// </summary>
        private static void UpdateCreatedDate()
        {
            //Method variables
            string createdColumnName = "created";

            //Get all data from the CSV file
            List<KnowledgePortal> knowledgePortal = KnowledgePortal.GetData();
            Console.WriteLine($"Source CSV file parsed with {knowledgePortal.Count} items.");

            //Get the SharePoint Site Id
            string siteId = _graphApi.GetSiteId(Configuration.SiteName);
            //Console.WriteLine($"Site ID: {siteId}");

            //Get the SharePoint List Id
            string listId = _graphApi.GetListId(siteId, Configuration.ListName);
            //Console.WriteLine($"List ID: {listId}");

            //Set the Created column to Read Only = False
            SetColumnReadOnly(siteId, listId, createdColumnName, false);

            //Get the all items in the list. Includes paging.
            List<SPListItem> listItems = _graphApi.GetListItems(siteId, listId);
            Console.WriteLine($"List {Configuration.ListName} has {listItems.Count} items with a Title.");

            //Loop through items and update the Created On value from the CSV file based on matching Title.
            foreach (var item in listItems)
            {
                //Match the title of the list item to the Name field in the CSV file.
                var matchItem = knowledgePortal.Where(a => a.Name.ToLower() == item.fields.Title.ToLower()).FirstOrDefault();

                //If there is a match, update the Created value based on the value in the CSV file
                if (matchItem != null)
                {
                    var createdOn = matchItem.Created;
                    _graphApi.UpdateListItem(siteId, listId, item.id, createdOn, null);
                    Console.WriteLine($"Updated Item with ID: {item.id} | Name: {item.fields.Title} | Created: {createdOn}");
                }
                else
                {
                    Console.WriteLine($"Skipped Item with ID: {item.id} | Name: {item.fields.Title} | Not found in CSV Data.");
                }
            }

            //Set the Created column back to Read Only = True
            SetColumnReadOnly(siteId, listId, createdColumnName, true);
        }

        /// <summary>
        /// Downloads all files in a Document Library and gets the PDF metadata from the PDF documents in the library
        /// </summary>
        private static void ExportAuthorInfo()
        {
            string siteId = _graphApi.GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string listId = _graphApi.GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            string driveId = _graphApi.GetDriveId(siteId);
            Console.WriteLine($"Drive ID: {driveId}");

            var driveItems = _graphApi.GetDriveItems(siteId, driveId);
            Console.WriteLine($"Retrieved Drive Items");

            var downloadedFiles = _graphApi.DownloadFiles(driveItems);
            Console.WriteLine($"Downloaded {downloadedFiles} files to local drive.");

            GetPdfFileInfo();
        }

        /// <summary>
        /// Update the author value in SharePoint
        /// </summary>
        private static void UpdateAuthor()
        {
            throw new NotImplementedException();

            string siteId = _graphApi.GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string driveId = _graphApi.GetDriveId(siteId);
            Console.WriteLine($"Drive ID: {driveId}");

            string listId = _graphApi.GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            var driveItems = _graphApi.GetDriveItems(siteId, driveId);
            Console.WriteLine($"Retrieved Drive Items");

            _graphApi.DownloadFiles(driveItems);
            GetPdfFileInfo();

            SPListDefinition listColumns = _graphApi.GetListDefinition(siteId, listId);
            SPColumnDefinition createdByColumn = new SPColumnDefinition();

            createdByColumn = _graphApi.GetColumnDefinition(listColumns, "author");
            Console.WriteLine($"Created By Column Id: {createdByColumn.id}");

            if (createdByColumn.readOnly)
            {
                _graphApi.SetColumnReadOnly(siteId, listId, createdByColumn.id, false);
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

            _graphApi.SetColumnReadOnly(siteId, listId, createdByColumn.id, true);
            Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag has been set to true");
        }

        #endregion Primary Methods

        #region Helper Methods

        /// <summary>
        /// Sets the read only flag of a column to true/false to update the column.
        /// </summary>
        /// <param name="siteId"></param>
        /// <param name="listId"></param>
        /// <param name="columnName"></param>
        /// <param name="targetValue"></param>
        private static void SetColumnReadOnly(string siteId, string listId, string columnName, bool targetValue)
        {
            //Get the definition of columns for the list
            SPListDefinition listColumns = _graphApi.GetListDefinition(siteId, listId);

            //Get a specific column definition from the list by the name
            SPColumnDefinition columnDefinition = _graphApi.GetColumnDefinition(listColumns, columnName);
            Console.WriteLine($"Created On Column Id: {columnDefinition.id}");

            //If the column is read only, set it to not read only.
            if (columnDefinition.readOnly == targetValue)
            {
                _graphApi.SetColumnReadOnly(siteId, listId, columnDefinition.id, targetValue);
                Console.WriteLine($"Column {columnName} ReadOnly flag has been set to {targetValue}");
            }
            else
            {
                Console.WriteLine($"Column {columnDefinition.name} ReadOnly flag is already {targetValue}");
            }
        }

        /// <summary>
        /// Get file info for PDF documents and export to a CSV file
        /// </summary>
        private static void GetPdfFileInfo()
        {
            string[] docNames = PdfDoc.GetPdfDocumentsFromDirectory(Configuration.DownloadDirectory);

            List<PdfDoc> pdfDocuments = new List<PdfDoc>();

            foreach (var doc in docNames)
            {
                PdfDoc pdf = new PdfDoc(doc);
                if (!pdf.Error)
                {
                    pdfDocuments.Add(pdf);
                }
            }

            string outputFile = PdfDoc.ExportPdfInfoToCsv(pdfDocuments);
            Console.WriteLine("Pdf Metadata Exported to {0} for {1} PDF Documents.", outputFile, pdfDocuments.Count);
        }

        #endregion Helper Methods




    }

    



    
}
