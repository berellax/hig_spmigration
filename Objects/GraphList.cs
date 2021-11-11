using System;
using System.Collections.Generic;
using System.Text;

namespace HIGKnowledgePortal
{
    public class GraphList
    {
        public List<GraphListItem> value { get; set; }
    }
    public class GraphListItem
    {
        public string id { get; set; }
        public string createdDateTime { get; set; }
        public string webUrl { get; set; }
        public User createdBy { get; set; }
        public Fields fields { get; set; }
        public class User
        {
            public string email { get; set; }
            public string id { get; set; }
            public string displayName { get; set; }
        }

        public class Fields
        {
            public string Title { get; set; }
            public string EMail { get; set; }
            public string Name { get; set; }
        }
    }
}
