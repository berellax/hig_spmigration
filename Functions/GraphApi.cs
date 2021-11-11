using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace HIGKnowledgePortal
{
    public class GraphApi
    {
        private readonly string _graphUrl;

        /// <summary>
        /// Constructor will authenticate the Graph API
        /// </summary>
        public GraphApi(AuthenticationType authType)
        {
            Configuration.AuthToken = AuthenticateGraphApi(authType);
            _graphUrl = Configuration.GraphUrl;
        }

        /// <summary>
        /// Authenticate to the Graph API based on the AuthType selected
        /// </summary>
        /// <returns></returns>
        private string AuthenticateGraphApi(AuthenticationType authType)
        {
            string url = $"https://login.windows.net/{Configuration.AuthTenant}/oauth2/token/";
            string postBody = string.Empty;

            if (authType == AuthenticationType.AppRegistration)
            {
                postBody = $"&client_id={Configuration.AuthClient}";
                postBody += @"&grant_type=client_credentials";
                postBody += $"&client_secret={Configuration.AuthClientSecret}";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default";
            }
            else if (authType == AuthenticationType.UserAuthentication)
            {
                Console.WriteLine("Username: ");
                var userName = Console.ReadLine();

                Console.WriteLine("Password: ");
                var password = Console.ReadLine();

                postBody = $"&client_id={Configuration.AuthClient}";
                postBody += $"&grant_type=password";
                postBody += @"&resource=https%3A%2F%2Fgraph.microsoft.com%2F";
                postBody += $"&username={userName}";
                postBody += $"&password={password}";
                postBody += $"&scope=openid";
            }
            else if (authType == AuthenticationType.AuthenticationToken)
            {
                Console.WriteLine("Paste Graph API Authentication Token Here:");
                var authToken = Console.ReadLine();
                return authToken;
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
                Console.WriteLine($"An error occurred authenticating your request. {ex.Message}.");
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
        public List<GraphListItem> GetListItems(string siteId, string listId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title)";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject responseJson = JsonConvert.DeserializeObject<JObject>(response);

            GraphList list = (GraphList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(GraphList));

            while (responseJson.ContainsKey("@odata.nextLink"))
            {
                request = WebFunctions.GetWebRequest(responseJson["@odata.nextLink"].ToString(), "GET");

                response = WebFunctions.GetWebResponse(request);

                responseJson = JsonConvert.DeserializeObject<JObject>(response);

                GraphList newList = (GraphList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(GraphList));

                list.value.AddRange(newList.value);
            }

            List<GraphListItem> listItems = list.value.Where(a => a.fields.Title != null).ToList();
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

        public JToken GetDriveItems(string siteId, string driveId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/drives/{driveId}/root/children";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            JObject json = (JObject)JsonConvert.DeserializeObject(response);

            if (!json.ContainsKey("value"))
            {
                return null;
            }

            JToken jsonValue = json["value"];

            while (json.ContainsKey("@odata.nextLink"))
            {
                request = WebFunctions.GetWebRequest(json["@odata.nextLink"].ToString(), "GET");

                response = WebFunctions.GetWebResponse(request);

                json = JsonConvert.DeserializeObject<JObject>(response);

                jsonValue.Append(json["value"]);
            }

            return jsonValue; ;
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

        public List<GraphListItem> GetUserListItems(string siteId, string listId)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items?expand=fields(select=EMail,Name)";

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "GET");

            string response = WebFunctions.GetWebResponse(request);

            GraphList list = (GraphList)System.Text.Json.JsonSerializer.Deserialize(response, typeof(GraphList));
            List<GraphListItem> listItems = list.value.Where(a => a.fields.EMail != null).ToList();
            return listItems;
        }

        public void UpdateListItem(string siteId, string listId, string itemId, JObject updateObject)
        {
            var url = $"{_graphUrl}/sites/{siteId}/lists/{listId}/items/{itemId}/fields";

            if(updateObject == null)
            {
                Console.WriteLine("Update values not specified.");
                return;
            }

            var content = JsonConvert.SerializeObject(updateObject);

            HttpWebRequest request = WebFunctions.GetWebRequest(url, "PATCH", content);
            string webResponse = WebFunctions.GetWebResponse(request);
        }

        public int DownloadFiles(JToken driveItems)
        {
            int i = 0;
            foreach (var item in driveItems)
            {
                if (item["file"] != null && item["name"] != null && item["name"].ToString().Contains(".pdf"))
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
            List<GraphListItem> userList = GetUserListItems(siteId, listId);

            GraphListItem createdByUser = userList.Where(a => a.fields.EMail.ToLower() == matchItem.CreatedBy.ToLower()).FirstOrDefault();

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

    public enum AuthenticationType
    {
        AppRegistration,
        UserAuthentication,
        AuthenticationToken
    }
}
