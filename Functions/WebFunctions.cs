using System.IO;
using System.Net;
using System.Text;

namespace HIGKnowledgePortal
{
    static class WebFunctions
    {
        public static HttpWebRequest GetWebRequest(string url, string method, string content = null, string contentType = "application/json")
        {
            var request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = method;
            request.ContentType = contentType;
            request.Accept = "application/json";
            request.Headers.Add("Authorization", "Bearer " + Configuration.AuthToken);

            if (content != null)
            {
                var data = Encoding.Default.GetBytes(content);
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                request.ContentLength = data.Length;
            }

            return request;
        }

        public static string GetWebResponse(HttpWebRequest request)
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
    }


}
