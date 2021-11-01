using Microsoft.Extensions.Configuration;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace KeebTalentBook
{
    public class Program
    {
        private static string FILE_PATH;
        private static string GRAPH_URL;
        private static string SP_ROOT;
        private static string SITE_NAME;
        private static string LIST_NAME;
        private static string AUTH_TOKEN;
        private static string AD_TENANT;
        private static string AD_CLIENT_ID;
        private static string AD_CLIENT_SECRET;
        private static string AD_USERNAME;
        private static string AD_PASSWORD;
        private static string AUTH_TYPE;
        private static bool SET_AUTHOR;
        const string USER_AUTH = "user_authentication";
        const string APP_AUTH = "app_registration";
        const string TOKEN_AUTH = "auth_token";

        static void Main(string[] args)
        {
            GetConfiguration();

            if(AUTH_TYPE != TOKEN_AUTH)
            {
                AuthenticateGraphApi();
            }

            List<KnowledgePortal> knowledgePortal = GetKnowledgePortalData();
            Console.WriteLine($"Text file parsed with {knowledgePortal.Count} items.");

            string siteId = GetSiteId();
            Console.WriteLine($"Site ID: {siteId}");

            string listId = GetListId(siteId);
            Console.WriteLine($"List ID: {listId}");

            SPListDefinition listColumns = GetListDefinition(siteId, listId);
            SPColumnDefinition createdByColumn = new SPColumnDefinition();

            if (SET_AUTHOR)
            {
                createdByColumn = GetColumnDefinition(listColumns, "author");
                Console.WriteLine($"Created By Column Id: {createdByColumn.id}");

                if (createdByColumn.readOnly)
                {
                    SetColumnReadOnly(siteId, listId, createdByColumn.id, "false");
                    Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag has been set to false");
                }
                else
                {
                    Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag is already false");
                }
            }


            SPColumnDefinition createdOnColumn = GetColumnDefinition(listColumns, "created");
            Console.WriteLine($"Created On Column Id: {createdOnColumn.id}");

            if (createdOnColumn.readOnly)
            {
                SetColumnReadOnly(siteId, listId, createdOnColumn.id, "false");
                Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag has been set to false");
            }
            else
            {
                Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag is already false");
            }

            List<SPListItem> listItems = GetListItems(siteId, listId);
            Console.WriteLine($"List {LIST_NAME} has {listItems.Count} items with a Title.");

            foreach (var item in listItems)
            {
                var matchItem = knowledgePortal.Where(a => a.Name == item.fields.Title).FirstOrDefault();
                if (matchItem != null)
                {
                    var createdOn = matchItem.Created;
                    string createdBy = null;

                    if (SET_AUTHOR)
                    {
                        matchItem.CreatedBy = "backup@keeebbob.onmicrosoft.com";
                        createdBy = GetCreatedByValue(matchItem, siteId);
                    }

                    UpdateListItem(siteId, listId, item.id, createdOn, createdBy);
                    Console.WriteLine($"Updated Item with ID: {item.id} | Name: {item.fields.Title} | Created: {createdOn}");
                }
            }

            SetColumnReadOnly(siteId, listId, createdOnColumn.id, "true");
            Console.WriteLine($"Column {createdOnColumn.name} ReadOnly flag has been set to true");

            if (SET_AUTHOR)
            {
                SetColumnReadOnly(siteId, listId, createdByColumn.id, "true");
                Console.WriteLine($"Column {createdByColumn.name} ReadOnly flag has been set to true");
            }
 

            Console.WriteLine("Application has completed. Press any key to exit");
            Console.ReadLine();
        }

        #region Graph API
        private static void AuthenticateGraphApi()
        {
            //var url = $"https://login.windows.net/{AD_TENANT}/oauth2/v2.0/authorize?";
            //url += $"client_id={AD_CLIENT_ID}";
            //url += @"&response_type=id_token+code";
            //url += @"&scope=opendid";
            //url += @"&response_mode=fragment";

            string url = $"https://login.windows.net/{AD_TENANT}/oauth2/token/";
            string postBody = string.Empty;

            if (AUTH_TYPE == "app_registration")
            {
                postBody = $"&client_id={AD_CLIENT_ID}";
                postBody += @"&grant_type=client_credentials";
                postBody += $"&client_secret={AD_CLIENT_SECRET}";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default";
            }
            else if (AUTH_TYPE == "user_authentication")
            {
                postBody = $"&client_id={AD_CLIENT_ID}";
                postBody += $"&grant_type=password";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += $"&username={AD_USERNAME}";
                postBody += $"&password={AD_PASSWORD}";
                postBody += $"&scope=openid";
            }
            else
            {
                Console.WriteLine("Authentication type not specified.");
                Environment.Exit(-1);
            }

            try
            {
                HttpWebRequest request = GetWebRequest(url, "POST", postBody, null);
                string response = GetWebResponse(request);
                var jsonResult = JObject.Parse(response);

                if (jsonResult.ContainsKey("access_token"))
                {
                    AUTH_TOKEN = (string)jsonResult["access_token"];
                }
                else
                {
                    throw new Exception("Access Token not returned");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred authenticating your request. {ex.Message}. Press any key to close.");
                Console.ReadLine();
                Environment.Exit(-1);
            }

        }

        private static string GetSiteId(string siteName = null)
        {
            if (siteName == null)
                siteName = SITE_NAME;

            var url = $"{GRAPH_URL}/sites/{SP_ROOT}:/sites/{siteName}";

            HttpWebRequest request = GetWebRequest(url, "GET");

            string response = GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            var siteId = json["id"].ToString();

            return siteId;

        }

        private static string GetListId(string siteId, string listName = null)
        {
            if (listName == null)
                listName = LIST_NAME;

            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listName}";

            HttpWebRequest request = GetWebRequest(url, "GET");

            string response = GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            string listId = json["id"].ToString();

            return listId;

        }

        private static SPListDefinition GetListDefinition(string siteId, string listId)
        {
            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listId}/columns";

            HttpWebRequest request = GetWebRequest(url, "GET");

            string response = GetWebResponse(request);

            SPListDefinition columns = (SPListDefinition)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPListDefinition));

            return columns;
        }

        private static SPColumnDefinition GetColumnDefinition(SPListDefinition listDefinition, string columnName)
        {
            SPColumnDefinition columnDefinition = listDefinition.value.Where(a => a.name.ToLower() == columnName).FirstOrDefault();

            return columnDefinition;
        }

        private static void SetColumnReadOnly(string siteId, string listId, string columnId, string targetValue)
        {
            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listId}/columns/{columnId}";

            JObject json = new JObject();
            json["readOnly"] = targetValue;

            var content = JsonConvert.SerializeObject(json);

            HttpWebRequest request = GetWebRequest(url, "PATCH", content);

            string response = GetWebResponse(request);
        }

        private static List<SPListItem> GetListItems(string siteId, string listId)
        {
            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title)";

            HttpWebRequest request = GetWebRequest(url, "GET");

            string response = GetWebResponse(request);

            SPList list = (SPList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPList));
            List<SPListItem> listItems = list.value.Where(a => a.fields.Title != null).ToList();
            return listItems;
        }

        private static List<SPListItem> GetUserListItems(string siteId, string listId)
        {
            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listId}/items?expand=fields(select=EMail,Name)";

            HttpWebRequest request = GetWebRequest(url, "GET");

            string response = GetWebResponse(request);

            SPList list = (SPList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPList));
            List<SPListItem> listItems = list.value.Where(a => a.fields.EMail != null).ToList();
            return listItems;

            //JObject list = (JObject)JsonConvert.DeserializeObject(response);

            //return list["value"];
        }

        private static string GetCreatedByValue(KnowledgePortal matchItem, string siteId)
        {
            var listId = GetListId(siteId, "User Information List");
            //var columns = GetListDefinition(siteId, listId);
            List<SPListItem> userList = GetUserListItems(siteId, listId);

            SPListItem createdByUser = userList.Where(a => a.fields.EMail.ToLower() == matchItem.CreatedBy.ToLower()).FirstOrDefault();
            
            if(createdByUser != null)
            {
                return createdByUser.id;
            }
            else
            {
                return null;
            }

            //var url = $"{GRAPH_URL}/users?$filter=mail eq '{matchItem.CreatedBy}'";

            //HttpWebRequest request = GetWebRequest(url, "GET");
            //string response = GetWebResponse(request);

            //JObject json = (JObject)JsonConvert.DeserializeObject(response);

            //var firstValue = json["value"].First();

            //return firstValue["id"].ToString();
        }


        private static void UpdateListItem(string siteId, string listId, string itemId, string createdOn, string createdBy)
        {
            var url = $"{GRAPH_URL}/sites/{siteId}/lists/{listId}/items/{itemId}/fields";

            JObject json = new JObject();
            json["Created"] = createdOn;

            if(SET_AUTHOR && createdBy != null)
                json["AuthorLookupId"] = createdBy;

            var content = JsonConvert.SerializeObject(json);

            HttpWebRequest request = GetWebRequest(url, "PATCH", content);
            string webResponse = GetWebResponse(request);
        }

        #endregion Graph API

        #region Helpers
        private static HttpWebRequest GetWebRequest(string url, string method, string content = null, string contentType = "application/json")
        {
            var request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = method;
            request.ContentType = contentType;
            request.Accept = "application/json";
            request.Headers.Add("Authorization", "Bearer " + AUTH_TOKEN);

            if (content != null)
            {
                var data = Encoding.Default.GetBytes(content);
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                request.ContentLength = data.Length;
            }

            return request;
        }

        private static string GetWebResponse(HttpWebRequest request)
        {
            HttpWebResponse response;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (WebException wex)
            {
                if (wex.Response == null)
                    return null;
                using (var errorResponse = (HttpWebResponse)wex.Response)
                {
                    using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            return new StreamReader(stream: response.GetResponseStream()).ReadToEnd();
        }

        private static void GetConfiguration()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appSettings.json")
                .Build();


            FILE_PATH = config.GetValue<string>("filePath");
            GRAPH_URL = config.GetValue<string>("graphUrl");
            SP_ROOT = config.GetValue<string>("spTenant");
            SITE_NAME = config.GetValue<string>("spSite");
            LIST_NAME = config.GetValue<string>("spList");
            SET_AUTHOR = config.GetValue<bool>("setAuthor");

            var auth = config.GetSection("authentication");
            AUTH_TYPE = auth.GetValue<string>("type");
            AD_TENANT = auth.GetValue<string>("tenant");
            AD_CLIENT_ID = auth.GetValue<string>("client_id");


            if (AUTH_TYPE == USER_AUTH)
            {
                AD_USERNAME = auth.GetValue<string>("username");
                AD_PASSWORD = auth.GetValue<string>("password");
            }
            else if (AUTH_TYPE == APP_AUTH)
            {
                AD_CLIENT_SECRET = auth.GetValue<string>("client_secret");
            }
            else if(AUTH_TYPE == TOKEN_AUTH)
            {
                AUTH_TOKEN = auth.GetValue<string>("auth_token");
            }
        }
        #endregion Helpers

        #region KnowledgePortal
        private static List<KnowledgePortal> GetKnowledgePortalData()
        {
            List<KnowledgePortal> knowledgePortals = new List<KnowledgePortal>();
            using (TextFieldParser parser = new TextFieldParser(FILE_PATH))
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

                    knowledgePortals.Add(CreateKnowledgePortal(row));
                }
            }

            return knowledgePortals;
        }

        private static KnowledgePortal CreateKnowledgePortal(string[] row)
        {
            KnowledgePortal knowledgePortal = new KnowledgePortal()
            {
                HomeOffice = row[0],
                DealStatus = row[1],
                Name = row[2],
                Created = row[3],
                Country = row[4],
                Industry = row[5],
                Keywords = row[6],
                Modified = row[7],
                Fund = row[8],
                DriveUrl = row[9],
                CompanyName = row[10],
                Key = row[11],
                DriveId = row[12],
                DocumentType = row[13],
                DealTeam = row[14],

            };

            return knowledgePortal;
        }
        #endregion KnowledgePortal

        #region TalentBook
        //private static List<TalentBook> GetTalentBookData()
        //{
        //    List<TalentBook> talentBooks = new List<TalentBook>();
        //    using (TextFieldParser parser = new TextFieldParser(FILE_PATH))
        //    {
        //        bool firstLine = true;
        //        string[] headers;
        //        parser.TextFieldType = FieldType.Delimited;
        //        parser.SetDelimiters(",");

        //        while (!parser.EndOfData)
        //        {
        //            string[] row = parser.ReadFields();

        //            if (firstLine)
        //            {
        //                headers = row;
        //                firstLine = false;
        //                continue;
        //            }

        //            talentBooks.Add(CreateTalentBook(row));
        //        }
        //    }

        //    return talentBooks;
        //}

        //private static TalentBook CreateTalentBook(string[] row)
        //{
        //    TalentBook talentBook = new TalentBook()
        //    {
        //        Status = row[0],
        //        Picture = row[1],
        //        HomeOffice = row[2],
        //        ModifiedBy = row[3],
        //        Name = row[4],
        //        Title = row[5],
        //        Language = row[6],
        //        CreatedOn = row[7],
        //        Degrees = row[8],
        //        Industry = row[9],
        //        ModifiedOn = row[10],
        //        Email = row[11],
        //        Fund = row[12],
        //        PriorExperience = row[13],
        //        OfficeExtension = row[14],
        //        StartDate = row[15],
        //        FunctionalExpertise = row[16]
        //    };

        //    return talentBook;
        //}
        #endregion TalentBook
    }

    internal class KnowledgePortal
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
    }

    //internal class TalentBook
    //{
    //    public string Status { get; set; }
    //    public string Picture { get; set; }
    //    public string HomeOffice { get; set; }
    //    public string ModifiedBy { get; set; }
    //    public string Name { get; set; }
    //    public string Title { get; set; }
    //    public string Language { get; set; }
    //    public string CreatedOn { get; set; }
    //    public string Degrees { get; set; }
    //    public string Industry { get; set; }
    //    public string ModifiedOn { get; set; }
    //    public string Email { get; set; }
    //    public string Fund { get; set; }
    //    public string PriorExperience { get; set; }
    //    public string OfficeExtension { get; set; }
    //    public string StartDate { get; set; }
    //    public string FunctionalExpertise { get; set; }

    //}

    internal class SPListDefinition
    {
        public List<SPColumnDefinition> value { get; set; }
    }

    internal class SPColumnDefinition
    {
        public string columnGroup { get; set; }
        public string description { get; set; }
        public string displayName { get; set; }
        public bool enforceUniqueValues { get; set; }
        public bool hidden { get; set; }
        public string id { get; set; }
        public bool indexed { get; set; }
        public string name { get; set; }
        public bool readOnly { get; set; }
        public bool required { get; set; }

    }

    internal class SPList
    {
        public List<SPListItem> value { get; set; }
    }
    internal class SPListItem
    {
        public string id { get; set; }
        public string createdDateTime { get; set; }
        public string webUrl { get; set; }
        public User createdBy { get; set; }
        public Fields fields { get; set; }

    }

    internal class User
    {
        public string email { get; set; }
        public string id { get; set; }
        public string displayName { get; set; }
    }

    internal class Fields
    {
        public string Title { get; set; }
        public string EMail { get; set; }
        public string Name { get; set; }
    }
}
