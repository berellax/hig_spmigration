﻿using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace KeebTalentBook
{
    public class GraphApi
    {
        const string USER_AUTH = "user_authentication";
        const string APP_REGISTRATION = "app_registration";
        const string TOKEN_AUTH = "auth_token";
        private readonly string _graphUrl;

        /// <summary>
        /// Constructor will authenticate the Graph API
        /// </summary>
        public GraphApi()
        {
            Configuration.AuthToken = AuthenticateGraphApi();
            _graphUrl = Configuration.GraphUrl;
        }

        /// <summary>
        /// Authenticate to the Graph API based on the AuthType selected
        /// </summary>
        /// <returns></returns>
        private string AuthenticateGraphApi()
        {
            string url = $"https://login.windows.net/{Configuration.AuthTenant}/oauth2/token/";
            string postBody = string.Empty;

            if (Configuration.AuthType == APP_REGISTRATION)
            {
                postBody = $"&client_id={Configuration.AuthClient}";
                postBody += @"&grant_type=client_credentials";
                postBody += $"&client_secret={Configuration.AuthClientSecret}";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default";
            }
            else if (Configuration.AuthType == USER_AUTH)
            {
                postBody = $"&client_id={Configuration.AuthClient}";
                postBody += $"&grant_type=password";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += $"&username={Configuration.AuthUsername}";
                postBody += $"&password={Configuration.AuthPassword}";
                postBody += $"&scope=openid";
            }
            else if (Configuration.AuthType == TOKEN_AUTH)
            {
                return Configuration.AuthToken;
            }
            else
            {
                Console.WriteLine("Authentication type not specified.");
                throw new Exception("Authentication Type Not Specified");
            }

            try
            {
                HttpWebRequest request = WebFunctions.GetWebRequest(url, "POST", postBody, null);
                string response = WebFunctions.GetWebResponse(request);
                var jsonResult = JObject.Parse(response);

                if (jsonResult.ContainsKey("access_token"))
                {
                    return (string)jsonResult["access_token"];
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
                throw new Exception("An error occurred authenticating your request");
            }
        }

        /// <summary>
        /// Get the Site Id
        /// </summary>
        /// <param name="siteName"></param>
        /// <returns></returns>
        public string GetSiteId(string siteName = null)
        {
            if (siteName == null)
                siteName = Configuration.SiteName;

            var url = $"{_graphUrl}/sites/{Configuration.SharePointOrg}:/sites/{Configuration.SiteName}";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            var siteId = json["id"].ToString();

            return siteId;      
        }

        /// <summary>
        /// Get the List Id for a List Name within a Site
        /// </summary>
        /// <param name="siteId"></param>
        /// <param name="listName"></param>
        /// <returns></returns>
        public string GetListId(string siteId, string listName = null)
        {
            if (listName == null)
                listName = Configuration.ListName;

            var url = $"{_graphUrl}/sites/{siteId}/lists/{listName}";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            string listId = json["id"].ToString();

            return listId;
        }

        /// <summary>
        /// Get a list of items in a SharePoint List
        /// </summary>
        /// <param name="siteId"></param>
        /// <param name="listId"></param>
        /// <returns></returns>
        public List<SPListItem> GetListItems(string siteId, string listId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title)";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject responseJson = JsonConvert.DeserializeObject<JObject>(response);

            SPList list = (SPList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPList));

            while (responseJson.ContainsKey("@odata.nextLink"))
            {
                request = WebFunctions.GetWebRequest(responseJson["@odata.nextLink"].ToString(), "GET");

                response = WebFunctions.GetWebResponse(request);

                responseJson = JsonConvert.DeserializeObject<JObject>(response);

                SPList newList = (SPList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPList));

                list.value.AddRange(newList.value);
            }

            List<SPListItem> listItems = list.value.Where(a => a.fields.Title != null).ToList();
            return listItems;
        }

        /// <summary>
        /// Get the Drive Id for a Drive Name within a Site
        /// </summary>
        /// <param name="siteId"></param>
        /// <param name="driveName"></param>
        /// <returns></returns>
        public string GetDriveId(string siteId, string driveName = null)
        {
            if (driveName == null)
                driveName = Configuration.DriveName;

            string driveId = null;

            var siteDrives = GetSiteDrives(siteId);

            if (siteDrives != null)
            {
                var drive = siteDrives.Where(d => d["name"].ToString() == driveName).FirstOrDefault();

                if (drive != null)
                {
                    driveId = drive["id"].ToString();
                }
            }

            return driveId;
        }
        private JToken GetSiteDrives(string siteId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/drives";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            JToken jsonValue = null;

            if (json.ContainsKey("value"))
            {
                jsonValue = json["value"];
            }

            return jsonValue;
        }

        public JObject GetDriveItems(string siteId, string driveId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/drives/{driveId}/root/children";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            if (!json.ContainsKey("value"))
            {
                return null;
            }

            return json; ;
        }

        public SPListDefinition GetListDefinition(string siteId, string listId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/columns";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            SPListDefinition columns = (SPListDefinition)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPListDefinition));

            return columns;
        }

        public SPColumnDefinition GetColumnDefinition(SPListDefinition listDefinition, string columnName)
        {
            SPColumnDefinition columnDefinition = listDefinition.value.Where(a => a.name.ToLower() == columnName).FirstOrDefault();

            return columnDefinition;
        }

        public void SetColumnReadOnly(string siteId, string listId, string columnId, bool targetValue)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/columns/{columnId}";

            JObject json = new JObject();
            json["readOnly"] = targetValue;

            var content = JsonConvert.SerializeObject(json);

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "PATCH", content);

            string response = WebFunctions.GetWebResponse(request);
        }

        public List<SPListItem> GetUserListItems(string siteId, string listId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items?expand=fields(select=EMail,Name)";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            SPList list = (SPList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(SPList));
            List<SPListItem> listItems = list.value.Where(a => a.fields.EMail != null).ToList();
            return listItems;
        }

        public void UpdateListItem(string siteId, string listId, string itemId, string createdOn, string createdBy)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items/{itemId}/fields";

            JObject json = new JObject();
            json["Created"] = createdOn;

            if (Configuration.SetAuthor && createdBy != null)
                json["AuthorLookupId"] = createdBy;

            var content = JsonConvert.SerializeObject(json);

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "PATCH", content);
            string webResponse = WebFunctions.GetWebResponse(request);
        }

        public int DownloadFiles(JObject driveItems)
        {
            int i = 0;
            foreach (JObject item in driveItems["value"])
            {
                if (item.ContainsKey("file"))
                {
                    var downloadUrl = item["@microsoft.graph.downloadUrl"].ToString();
                    var fileName = item["name"].ToString();
                    var success = FileDownloader.DownloadFile(downloadUrl, $"C:\\Temp\\{fileName}", 10000);
                    if (success)
                    {
                        i++;
                    }
                }
            }

            return i;
        }

        public string GetCreatedByValue(KnowledgePortal matchItem, string siteId)
        {
            var listId = GetListId(siteId, "User Information List");
            //var columns = GetListDefinition(siteId, listId);
            List<SPListItem> userList = GetUserListItems(siteId, listId);

            SPListItem createdByUser = userList.Where(a => a.fields.EMail.ToLower() == matchItem.CreatedBy.ToLower()).FirstOrDefault();

            if (createdByUser != null)
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

    }
}