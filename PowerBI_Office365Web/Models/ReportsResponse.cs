using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PowerBI_Office365Web.Models
{
    public class ReportsResponseValue
    {
        public string id { get; set; }
        public string name { get; set; }
        public string embedUrl { get; set; }
        public bool isReadOnly { get; set; }
    }

    public class ReportsResponse
    {
        [JsonProperty(PropertyName = "__@odata.context")]
        public string context { get; set; }
        public List<ReportsResponseValue> value { get; set; }
    }
}