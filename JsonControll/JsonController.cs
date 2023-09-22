using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TrainingAPI.App_Code;

namespace TrainingAPI.JsonControll
{
    static class JsonController
    {
        #region Parse Json
        public static Hashtable setJsonToParam(dynamic postData)
        {
            Hashtable param = null;
            param = new Hashtable();
            JToken jToken;
            dynamic parse = JArray.Parse(postData.ToString());
            dynamic array = parse[0];
            JObject obj = JObject.Parse(array.ToString());

            if (obj.Count > 0)
            {
                foreach (KeyValuePair<string, JToken> item in obj)
                {
                    //menerima parameter key and value
                    string key = item.Key;
                    jToken = item.Value;

                    // insert into param
                    param[key.ToString()] = jToken;

                }
            }
            return new Hashtable(param);
        }

        public static Hashtable setJsonToParams(dynamic postData)
        {
            Hashtable param = null;
            param = new Hashtable();
            JToken jToken;
            dynamic parse = JArray.Parse(postData.ToString());
            dynamic array = parse[0];
            JObject obj = JObject.Parse(array.ToString());

            if (obj.Count > 0)
            {
                foreach (KeyValuePair<string, JToken> item in obj)
                {
                    //menerima parameter key and value
                    string key = "p_" + item.Key;
                    jToken = item.Value;

                    // insert into param
                    param[key.ToString()] = jToken;

                }
            }
            return new Hashtable(param);
        }
        #endregion
    }

}
