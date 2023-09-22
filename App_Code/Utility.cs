using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Claims;
using Microsoft.AspNetCore.Identity;
using System.Net;
using Microsoft.Extensions.Primitives;

namespace TrainingAPI.App_Code
{
    public class Utility
    {
        protected IHttpContextAccessor _httpContextAccessor;

        public Utility(IHttpContextAccessor context)
        {
            _httpContextAccessor = context;
        }

        #region Convertion
        public static DateTime ToDateTime(string dt)
        {
            System.Globalization.DateTimeFormatInfo dtfi = null;

            try
            {
                dtfi = new System.Globalization.DateTimeFormatInfo();
                dtfi.ShortDatePattern = "dd/MM/yyyy";

                return DateTime.Parse(dt, dtfi);

            }
            catch (Exception)
            {
                return new DateTime(1900, 1, 1);
            }
        }
        #endregion

        #region Message Box
        public static string SAVE_DATA_SUCCESS_MESSAGE = "Success to save data.";
        public static string SAVE_DATA_FAIL_MESSAGE = "Fail to save data.";
        public static string LOAD_DATA_FAIL_MESSAGE = "Fail to load data.";
        public static string DELETE_VALIDATION_MESSAGE = "Are you sure to delete this record?";
        public static string DELETE_VALIDATION_SUCCESS_MESSAGE = "Success to delete data.";
        public static string DELETE_VALIDATION_FAIL_MESSAGE = "Fail to delete data.";

        public static void ApplyDefaultProp(Hashtable ht, string username, string ipaddress)
        {
            IPHostEntry host;
            //string localIP = "";

            //host = Dns.GetHostEntry(Dns.GetHostName());
            //localIP = host.AddressList[1].ToString();


            ht["p_cre_date"] = DateTime.Now;
            ht["p_cre_by"] = username;
            ht["p_cre_ip_address"] = ipaddress;

            ht["p_mod_date"] = DateTime.Now;
            ht["p_mod_by"] = username;
            ht["p_mod_ip_address"] = ipaddress;
        }
        #endregion
    }
}
