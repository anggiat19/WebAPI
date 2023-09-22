using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TrainingAPI.Models
{
    public class SearchCriteria
    {
        [JsonProperty(PropertyName = "filter")]
        public string Filter { get; set; }

        [JsonProperty(PropertyName = "isPageLoad")]
        public bool IsPageLoad { get; set; }
    }
}
