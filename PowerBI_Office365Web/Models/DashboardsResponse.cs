using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PowerBI_Office365Web.Models
{
    public class DashboardsResponseValue
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public bool isReadOnly { get; set; }
    }

    public class DashboardsResponse
    {
        [JsonProperty(PropertyName = "__@odata.context")]
        public string context{ get; set; }
        public List<DashboardsResponseValue> value { get; set; }
    }
}