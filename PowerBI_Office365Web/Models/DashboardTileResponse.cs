using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PowerBI_Office365Web.Models
{

    public class DashboardTileResponseValue
    {
        public string id { get; set; }
        public string title { get; set; }
        public string embedUrl { get; set; }
    }

    public class DashboardTileResponse
    {
        [JsonProperty(PropertyName = "__@odata.context")]
        public string context { get; set; }
        public List<DashboardTileResponseValue> value { get; set; }
    }
}