using System;
using System.Collections.Generic;
using System.Text;

namespace HIGKnowledgePortal
{
    public class SPListDefinition
    {
        public List<SPColumnDefinition> value { get; set; }
    }

    public class SPColumnDefinition
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
}
