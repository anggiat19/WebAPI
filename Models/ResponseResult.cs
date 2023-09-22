using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace TrainingAPI.Models
{
    public class ResponseResult
    {
        public class ResponseSuccessResult
        {
            [JsonProperty("result")]
            public int Result { get; set; } = 1;

            [JsonProperty("StatusCode")]
            public HttpStatusCode StatusCode { get; set; }

            [JsonProperty("message")]
            public string Message { get; set; } = "OK";

            [JsonProperty("data")]
            public object Data { get; set; } = "";

            [JsonProperty("code")]
            public string Code { get; set; } = "";

            [JsonProperty("id")]
            public Int64 Id { get; set; } = 0;

        }
        public class ResponseFailedResult
        {
            [JsonProperty("result")]
            public int Result { get; set; } = 0;

            [JsonProperty("StatusCode")]
            public HttpStatusCode StatusCode { get; set; }

            [JsonProperty("message")]
            public string Message { get; set; } = "ERROR";

            [JsonProperty("data")]
            public object Data { get; set; } = "";
        }
    }

}
