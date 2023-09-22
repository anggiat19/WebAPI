using System;
using System.Data;
using System.Collections;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using DataAccessLayer;
using TrainingAPI.JsonControll;
using TrainingAPI.App_Code;
using System.Threading.Tasks;
using static TrainingAPI.Models.ResponseResult;
using ExcelDataReader;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using TrainingAPI.Models;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Hosting;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace TrainingAPI.App_Code
{
    public class iFinancingController : ControllerBase
    {
        protected readonly string _connectionString;
        protected String SpNames = "";
        protected String TableNames = "";
        protected ResponseSuccessResult responseSuccessResult = new ResponseSuccessResult();
        protected ResponseFailedResult responseFailedResult = new ResponseFailedResult();
        protected readonly IWebHostEnvironment _hostingEnvironment;
        private String ConnectionString = "Server=imtec.ddns.net,7575; Database=IFINSYS; User ID=sa; Password=p@ssw0rd; Connection Timeout=0; Persist Security Info=true";
        private String RegexFilePathVal = @"(\S+)\.\.(\S+)";
        private String RegexFileExtension = @"^.*\.(png|PNG|jpg|JPG|jpeg|doc|docx|DOC|pdf|PDF|xls|xlsx|pptx|odt|ods|odp|zip|7z|rar|txt)$";
        private String secretKey = "PT Inovasi Mitra Sejati";
        protected string _docPrintSetting;
        protected String _docPathIfinsys;
        protected String _ifindocUploadUrl;
        protected String _ifindocDeleteUrl;

        //decodeHeader(Request.Headers["UserID"])
        //decodeHeader["IPAddress"]; // get IPAddress

        public iFinancingController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment)
        {
            //_connectionString = ConnectionString;
            _connectionString = configuration.GetConnectionString("DefaultConnection");
            _hostingEnvironment = hostingEnvironment;
            _docPathIfinsys = configuration.GetValue<string>("UrlApiCall:docPathIfinsys");
            _docPrintSetting = configuration.GetValue<string>("DocumentSetting:docPrintSetting");
            _ifindocUploadUrl = configuration.GetValue<string>("UrlApiCall:ifindocUploadUrl");
            _ifindocDeleteUrl = configuration.GetValue<string>("UrlApiCall:ifindocDeleteUrl");
        }

        #region getrows
        public object GetRows([FromBody]dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            int recordsTotal = 0;

            var pagingResponse = new PagingResponse()
            {
                Draw = data.draw
            };

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();
                var pageList = data.start / data.length + 1;
                int orderBy = data.order[0]["column"] - 1;
                string sortBy = data.order[0]["dir"];


                TableNames = string.Format("xsp_{0}_getrows", TableName);

                _htParameters["p_keywords"] = data.search.value;
                _htParameters["p_pagenumber"] = pageList;
                _htParameters["p_rowspage"] = data.length;
                _htParameters["p_order_by"] = orderBy;
                _htParameters["p_sort_by"] = sortBy;

                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());


                _dt = _dal.GetRows("", TableNames, _htParameters);

                if (_dt.Rows.Count != 0)
                {
                    recordsTotal = Convert.ToInt32(_dt.Rows[0]["rowcount"]);
                }

                pagingResponse.Province = _dt;
                pagingResponse.RecordsTotal = recordsTotal;
                pagingResponse.RecordsFiltered = recordsTotal;
            }
            catch (Exception ex)
            {
                pagingResponse.Msg = ex.InnerException.Message.ToString();
            }

            return pagingResponse;
        }

        public object GetRows([FromBody]dynamic data, string TableName, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            int recordsTotal = 0;

            var pagingResponse = new PagingResponse()
            {
                Draw = data.draw
            };

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();
                var pageList = data.start / data.length + 1;
                int orderBy = data.order[0]["column"] - 1;
                string sortBy = data.order[0]["dir"];

                // ini untuk dynamic getrows bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.paramTamp.ToString());
                // ini untuk dynamic sp name
                SpNames = SpName;

                _htParameters["p_keywords"] = data.search.value;
                _htParameters["p_pagenumber"] = pageList;
                _htParameters["p_rowspage"] = data.length;
                _htParameters["p_order_by"] = orderBy;
                _htParameters["p_sort_by"] = sortBy;
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());


                _dt = _dal.GetRows("", SpNames, _htParameters);

                if (_dt.Rows.Count != 0)
                {
                    recordsTotal = Convert.ToInt32(_dt.Rows[0]["rowcount"]);
                }

                pagingResponse.Province = _dt;
                pagingResponse.RecordsTotal = recordsTotal;
                pagingResponse.RecordsFiltered = recordsTotal;
            }
            catch (Exception ex)
            {
                pagingResponse.Msg = ex.InnerException.Message.ToString();
            }

            return pagingResponse;
        }
        #endregion

        #region getrow
        public object GetRow([FromBody]dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();


                // ini untuk dynamic getrow bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.ToString());
                // ini untuk dynamic table name
                //SpName = data[0].sp_name
                TableNames = string.Format("xsp_{0}_getrow", TableName);
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                _dt = _dal.GetRow("", TableNames, _htParameters);

                responseSuccessResult.Data = _dt;
                result = responseSuccessResult;
            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }

        public object GetRow([FromBody]dynamic data, string TableName, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();


                // ini untuk dynamic getrow bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.ToString());
                SpNames = SpName;
                // ini untuk dynamic table name
                //SpName = data[0].sp_name
                TableNames = string.Format("xsp_{0}_getrow", TableName);
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                _dt = _dal.GetRow("", SpNames, _htParameters);

                responseSuccessResult.Data = _dt;
                result = responseSuccessResult;
            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }
        #endregion getrow

        #region insert
        public object Insert([FromBody] dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname 
                    TableNames = string.Format("xsp_{0}_insert", TableName);
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Insert("", TableNames, _htParameters);
                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        public object Insert([FromBody] dynamic data, string TableName, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname 
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Insert("", SpName, _htParameters);
                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        public object InsertOutputCode([FromBody]dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;
            string code = "";

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                // ini untuk dynamic insert bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.ToString());
                // ini untuk dynamic spname
                TableNames = string.Format("xsp_{0}_insert", TableName);
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                _dal.Insert("", TableNames, _htParameters, ref code);

                responseSuccessResult.Data = _dal;
                responseSuccessResult.Code = code;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }

        public object InsertOutputId([FromBody]dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;
            long id = 0;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic update bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname
                    TableNames = string.Format("xsp_{0}_insert", TableName);
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _dal.Insert("", TableNames, _htParameters, ref id);
                }
                responseSuccessResult.Data = _dal;
                responseSuccessResult.Id = id;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }
        #endregion

        #region update
        public object Update([FromBody] dynamic data, string TableName, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic update bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Update("", SpName, _htParameters);
                }

                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        public object Update([FromBody] dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic update bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname
                    TableNames = string.Format("xsp_{0}_update", TableName);
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Update("", TableNames, _htParameters);
                }

                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion update

        #region delete
        public object Delete([FromBody] dynamic data, string TableName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);
                    // ini untuk dynamic spname
                    //SpName = data[0].sp_name;
                    TableNames = string.Format("xsp_{0}_delete", TableName);
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Delete("", TableNames, _htParameters);
                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion

        #region ExecSP
        public object ExecSp([FromBody]dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            object result = null;
            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                // ini untuk dynamic insert bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.ToString());
                if (data[0].action == "getResponse")
                {
                    // ini untuk dynamic spname
                    //SpName = data[0].sp_name;
                    SpNames = SpName;
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _dt = _dal.GetRows("", SpNames, _htParameters);

                    responseSuccessResult.Data = _dt;
                    result = responseSuccessResult;
                }
                else
                {
                    // ini untuk dynamic spname
                    //SpName = data[0].sp_name;
                    SpNames = SpName;
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.ExecRawSP(SpNames, _htParameters);

                    responseSuccessResult.Data = _dal;
                    result = responseSuccessResult;
                }
            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }
        #endregion

        #region image upload, priview, dan delete
        protected object PostUploadImage([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            string filePath = String.Empty;
            string filePathCode = String.Empty;
            string fileName = String.Empty;
            object result = null;
            string uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;
            string module;
            string header;
            string child;
            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                //membuat folder
                var datetime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);


                    _htParameters["p_file_path"] = "";
                    _htParameters["p_file_name"] = "";

                    _htParameters = JsonController.setJsonToParam(jsonConvert);

                    Utility.ApplyDefaultProp(_htParameters, Username, IPAddress);

                    //mengambil data dari parameter yang dikirim
                    _htParameters["p_acess_type"] = "UPLOAD";
                    filePathCode = _htParameters["p_file_path"].ToString();
                    fileName = filePathCode + "_" + datetime + "_" + _htParameters["p_file_name"].ToString();
                    module = _htParameters["p_module"].ToString();
                    header = _htParameters["p_header"].ToString();
                    child = _htParameters["p_child"].ToString();
                    module = uploadPath + "\\" + module;
                    header = module + "\\" + header;
                    child = header + "\\" + child + "\\";

                    filePath = string.Format("{0}{1}", child, fileName);

                    Match matchPath = Regex.Match(filePath, RegexFilePathVal, RegexOptions.IgnoreCase);

                    if (!matchPath.Success)
                    {
                        Match matchExt = Regex.Match(filePath, RegexFileExtension, RegexOptions.IgnoreCase);
                        if (matchExt.Success)
                        {
                            if (_docPrintSetting.Equals("0"))// untuk fisik
                            {
                                if (!Directory.Exists(module))
                                {
                                    Directory.CreateDirectory(module);
                                }
                                if (!Directory.Exists(header))
                                {
                                    Directory.CreateDirectory(header);
                                }
                                if (!Directory.Exists(child))
                                {
                                    Directory.CreateDirectory(child);
                                }
                                var fileZip = string.Format("{0}{1}", child, Path.GetFileNameWithoutExtension(fileName));

                                _htParameters["p_file_paths"] = (fileZip ?? "");
                                _htParameters["p_file_name"] = (fileName ?? "");

                                var base64 = _htParameters["p_base64"].ToString();

                                byte[] byteFile = Convert.FromBase64String(base64);

                                System.IO.File.WriteAllBytes(filePath, byteFile);
                                uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;

                                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                                {
                                    zip.Password = secretKey;
                                    zip.AddFile(filePath, "");
                                    zip.Comment = ("This zip file was created from IFINANCING at " + DateTime.Now.ToString("G"));
                                    zip.Save(fileZip);
                                }

                                if (System.IO.File.Exists(filePath))
                                {
                                    System.IO.File.Delete(filePath);

                                    _dal.Update("", SpName, _htParameters);

                                    _dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);
                                }
                            }
                            else
                            {
                                filePath = _htParameters["p_module"].ToString() + "\\" + _htParameters["p_header"].ToString() + "\\" + _htParameters["p_child"].ToString();

                                _htParameters["p_file_paths"] = (filePath ?? "");
                                _htParameters["p_file_name"] = (fileName ?? "");

                                _dal.Insert("FILE_" + _htParameters["p_module"].ToString() + '_' + _htParameters["p_header"].ToString(), _htParameters);

                                _dal.Update("", SpName, _htParameters);

                                _dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);
                            }
                        }
                        else
                        {
                            return Ok(new { result = 0, filepath = filePath, filename = fileName, data = "E;warning;File Extension Not Valid" });
                        }
                    }
                    else
                    {
                        return Ok(new { result = 0, filepath = filePath, filename = fileName, data = "E;warning;Path Not Valid" });
                    }
                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;
            }
            catch (IOException ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        protected async Task<object> PostUploadImage([FromBody]dynamic data, string SpName, IHttpClientFactory _httpClientFactory)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            string filePath = String.Empty;
            string filePathCode = String.Empty;
            string fileName = String.Empty;
            string uploadResult = String.Empty;
            string uploadFailedMessage = String.Empty;
            object result = null;
            var res = (string)null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                string tokenHeader = Request.Headers["Authorization"]; // get token from header
                string usernameHeader = Request.Headers["UserID"]; // get userid from header
                string ipaddressHeader = Request.Headers["IPAddress"]; // get ipaddress from header

                #region untuk by pass certificate
                //var spHandler = new HttpClientHandler()
                //{
                //    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) =>
                //    {
                //        return true;
                //    }
                //};
                #endregion untuk by pass certificate

                var request = new HttpRequestMessage(HttpMethod.Post, new Uri(_ifindocUploadUrl));
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Authorization", tokenHeader);
                request.Headers.Add("UserID", usernameHeader);
                request.Headers.Add("IPAddress", ipaddressHeader);

                var Jsondata = JsonConvert.SerializeObject(data);
                request.Content = new StringContent(Jsondata, Encoding.UTF8, "application/json");

                #region untuk by pass certificate
                //var client = new HttpClient(spHandler);
                #endregion untuk by pass certificate 

                var client = _httpClientFactory.CreateClient();

                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    res = await response.Content.ReadAsStringAsync();
                    responseSuccessResult.StatusCode = response.StatusCode;
                    responseSuccessResult.Data = JsonConvert.DeserializeObject<object>(res);

                    uploadResult = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("result").ToString();
                    fileName = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("filename").ToString();
                    filePath = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("filepath").ToString();
                    uploadFailedMessage = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("data").ToString();

                    if (uploadResult.Equals("1"))
                    {
                        //membuat folder
                        var datetime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        dynamic parse = JArray.Parse(data.ToString());

                        for (int i = 0; i < parse.Count; i++)
                        {
                            var convert = JsonConvert.SerializeObject(parse[i]);

                            string[] stringArray = new string[] { convert };
                            var jsonConvert = JsonConvert.SerializeObject(stringArray);


                            _htParameters["p_file_path"] = "";
                            _htParameters["p_file_name"] = "";

                            _htParameters = JsonController.setJsonToParam(jsonConvert);

                            // ini untuk dynamic spname
                            SpNames = SpName;

                            //mengambil data dari parameter yang dikirim 

                            _htParameters["p_file_paths"] = (filePath ?? "");
                            _htParameters["p_file_name"] = (fileName ?? "");

                            Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                            _dal.Update("", SpNames, _htParameters);
                        }
                        responseSuccessResult.Data = _dal;
                        result = responseSuccessResult;
                    }
                    else
                    {
                        responseFailedResult.StatusCode = response.StatusCode;
                        responseFailedResult.Message = uploadFailedMessage;
                        result = responseFailedResult;
                    }
                }
                else
                {
                    responseFailedResult.StatusCode = response.StatusCode;
                    responseFailedResult.Message = response.ReasonPhrase.ToString();
                    result = responseFailedResult;
                }
            }
            catch (IOException ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        protected object Priview([FromBody] dynamic data)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;
            string file_name = "";
            string file_path = "";
            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                var convert = JsonConvert.SerializeObject(parse[0]);
                string[] stringArray = new string[] { convert };

                var jsonconvert = JsonConvert.SerializeObject(stringArray);
                _htParameters = JsonController.setJsonToParam(jsonconvert);

                file_name = _htParameters["p_file_name"].ToString();
                file_path = _htParameters["p_file_paths"].ToString();
                Match match = Regex.Match(file_path, RegexFilePathVal, RegexOptions.IgnoreCase);
                if (!match.Success)
                {
                    Match matchExt = Regex.Match(file_path, RegexFileExtension, RegexOptions.IgnoreCase);
                    if (matchExt.Success)
                    {

                        byte[] b = System.IO.File.ReadAllBytes(file_path);
                        var base64 = Convert.ToBase64String(b);

                        return Ok(new { data = Convert.ToBase64String(b), filename = file_name });
                    }
                    else
                    {
                        return Ok(new { result = 0, message = "E;warning;File Extension Not Valid" });
                    }
                }
                else
                {
                    return Ok(new { result = 0, message = "E;warning;Path Not Valid" });
                }

            }
            catch (IOException e)
            {
                return Ok(new { result = 0, message = e.InnerException.Message });
            }
        }

        protected object Priview([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            string uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;
            string file_name = "";
            string file_path = "";

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                var convert = JsonConvert.SerializeObject(parse[0]);
                string[] stringArray = new string[] { convert };

                var jsonconvert = JsonConvert.SerializeObject(stringArray);
                _htParameters = JsonController.setJsonToParam(jsonconvert);

                Utility.ApplyDefaultProp(_htParameters, Username, IPAddress);

                file_name = _htParameters["p_file_name"].ToString();
                file_path = _htParameters["p_file_paths"].ToString();

                if (_docPrintSetting.Equals("0"))// untuk fisik
                {
                    int start = uploadPath.Length + 1;
                    string subStrFilePath = file_path.Substring(start);
                    var filePath = subStrFilePath.Split("\\");

                    _htParameters["p_module"] = filePath[0];
                    _htParameters["p_header"] = filePath[1];
                    _htParameters["p_child"] = filePath[2];
                    _htParameters["p_acess_type"] = "PREVIEW";

                    Match match = Regex.Match(file_path + Path.GetExtension(file_name), RegexFilePathVal, RegexOptions.IgnoreCase);
                    if (!match.Success)
                    {
                        Match matchExt = Regex.Match(file_path + Path.GetExtension(file_name), RegexFileExtension, RegexOptions.IgnoreCase);
                        if (matchExt.Success)
                        {

                            string source = file_path;
                            string target = _hostingEnvironment.ContentRootPath + _docPathIfinsys + "\\TEMP\\EXTRACT\\";
                            string searchFile = file_name;

                            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                            using (Ionic.Zip.ZipFile zip = Ionic.Zip.ZipFile.Read(source))
                            {
                                bool results = zip.ContainsEntry(searchFile);
                                if (results)
                                {
                                    zip.Password = secretKey;
                                    zip.ExtractAll(target, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                                }
                            }
                            byte[] b = System.IO.File.ReadAllBytes(target + file_name);
                            var base64 = Convert.ToBase64String(b);

                            if (System.IO.File.Exists(target + file_name))
                            {
                                System.IO.File.Delete(target + file_name);

                                //_dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);
                            }

                            return Ok(new { data = base64, filename = file_name });
                        }
                        else
                        {
                            return Ok(new { result = 0, data = "E;warning;File Extension Not Valid" });
                        }
                    }
                    else
                    {
                        return Ok(new { result = 0, data = "E;warning;Path Not Valid" });
                    }
                }
                else
                {
                    //int start = file_path.Length + 1;
                    //string subStrFilePath = file_path.Substring(start);
                    var filePath = file_path.Split("\\");

                    _htParameters["p_module"] = filePath[0];
                    _htParameters["p_header"] = filePath[1];
                    _htParameters["p_child"] = filePath[2];
                    _htParameters["p_doc_no"] = filePath[2];
                    _htParameters["p_acess_type"] = "PREVIEW";

                    _dt = _dal.GetRow("", SpName, _htParameters);

                    var base64 = Encoding.UTF8.GetString(Convert.FromBase64String(_dt.Rows[0].ItemArray[5].ToString()));

                    _dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);

                    return Ok(new { data = base64, filename = file_name });
                }
            }
            catch (IOException e)
            {
                return Ok(new { result = 0, data = e.InnerException.Message });
            }
        }

        protected object Deletefile([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;
            string uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;
            string file_name = "";
            string file_path = "";
            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);
                    string[] stringArray = new string[] { convert };

                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParam(jsonConvert);

                    Utility.ApplyDefaultProp(_htParameters, Username, IPAddress);

                    file_name = _htParameters["p_file_name"].ToString();
                    file_path = _htParameters["p_file_paths"].ToString();

                    if (_docPrintSetting.Equals("0"))// untuk fisik
                    {

                        if (file_path != null || file_path != string.Empty)
                        {

                            Match matchPath = Regex.Match(file_path + Path.GetExtension(file_name), RegexFilePathVal, RegexOptions.IgnoreCase);

                            if (!matchPath.Success)
                            {
                                Match matchExt = Regex.Match(file_path + Path.GetExtension(file_name), RegexFileExtension, RegexOptions.IgnoreCase);
                                if (matchExt.Success)
                                {
                                    if (System.IO.File.Exists(file_path))
                                    {
                                        System.IO.File.Delete(file_path);

                                        _htParameters["p_file_name"] = "";
                                        _htParameters["p_file_paths"] = "";

                                        _dal.Update("", SpName, _htParameters);

                                        int start = uploadPath.Length + 1;
                                        string subStrFilePath = file_path.Substring(start);
                                        var filePath = subStrFilePath.Split("\\");

                                        _htParameters["p_file_name"] = file_name;
                                        _htParameters["p_file_paths"] = file_path;
                                        _htParameters["p_module"] = filePath[0];
                                        _htParameters["p_header"] = filePath[1];
                                        _htParameters["p_child"] = filePath[2];
                                        _htParameters["p_acess_type"] = "DELETE";

                                        _dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);
                                    }
                                }

                            }
                            else
                            {
                                return Ok(new { result = 0, data = "E;warning;File Extension Not Valid" });
                            }

                        }
                        else
                        {
                            return Ok(new { result = 0, data = "E;warning;Path Not Valid" });
                        }

                    }

                    else
                    {
                        int start = uploadPath.Length + 1;
                        string subStrFilePath = file_path.Substring(start);
                        var filePath = subStrFilePath.Split("\\");

                        _htParameters["p_file_name"] = "";
                        _htParameters["p_file_paths"] = "";

                        _dal.Delete("FILE_" + filePath[0] + '_' + filePath[1], _htParameters);

                        _dal.Update("", SpName, _htParameters);

                        _htParameters["p_file_name"] = file_name;
                        _htParameters["p_file_paths"] = file_path;
                        _htParameters["p_module"] = filePath[0];
                        _htParameters["p_header"] = filePath[1];
                        _htParameters["p_child"] = filePath[2];
                        _htParameters["p_doc_no"] = filePath[2];
                        _htParameters["p_acess_type"] = "DELETE";

                        _dal.Insert("", "XSP_SYS_DOC_ACCESS_LOG_INSERT", _htParameters);

                    }

                    responseSuccessResult.Data = _dal;
                    result = responseSuccessResult;
                }
            }
            catch (IOException ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }

        protected async Task<object> Deletefile([FromBody] dynamic data, string SpName, IHttpClientFactory _httpClientFactory)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;
            string filePath = String.Empty;
            string file_path = "";
            string uploadResult = String.Empty;
            string uploadFailedMessage = String.Empty;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                var res = (string)null;
                string tokenHeader = Request.Headers["Authorization"]; // get token from header
                string usernameHeader = Request.Headers["UserID"]; // get userid from header
                string ipaddressHeader = Request.Headers["IPAddress"]; // get ipaddress from header

                #region untuk by pass certificate
                //var spHandler = new HttpClientHandler()
                //{
                //    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) =>
                //    {
                //        return true;
                //    }
                //};
                #endregion untuk by pass certificate

                var request = new HttpRequestMessage(HttpMethod.Post, new Uri(_ifindocDeleteUrl));
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Authorization", tokenHeader);
                request.Headers.Add("UserID", usernameHeader);
                request.Headers.Add("IPAddress", ipaddressHeader);

                var Jsondata = JsonConvert.SerializeObject(data);
                request.Content = new StringContent(Jsondata, Encoding.UTF8, "application/json");

                #region untuk by pass certificate
                //var client = new HttpClient(spHandler);
                #endregion untuk by pass certificate 

                var client = _httpClientFactory.CreateClient();

                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    res = await response.Content.ReadAsStringAsync();
                    responseSuccessResult.StatusCode = response.StatusCode;
                    responseSuccessResult.Data = JsonConvert.DeserializeObject(res);

                    uploadResult = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("result").ToString();
                    uploadFailedMessage = JObject.Parse(JObject.Parse(res).GetValue("value").ToString()).GetValue("data").ToString();

                    if (uploadResult.Equals("1"))
                    {
                        dynamic parse = JArray.Parse(data.ToString());
                        for (int i = 0; i < parse.Count; i++)
                        {
                            var convert = JsonConvert.SerializeObject(parse[i]);
                            string[] stringArray = new string[] { convert };

                            var jsonConvert = JsonConvert.SerializeObject(stringArray);

                            // ini untuk dynamic bisa menerima parameter lebih dari satu dari screen
                            _htParameters = JsonController.setJsonToParam(jsonConvert);

                            // ini untuk dynamic spname
                            SpNames = SpName;

                            _htParameters["p_file_name"] = "";
                            _htParameters["p_file_paths"] = "";

                            _dal.Update("", SpNames, _htParameters);
                        }
                        responseSuccessResult.Data = _dal;
                        result = responseSuccessResult;
                    }
                    else
                    {
                        responseFailedResult.StatusCode = response.StatusCode;
                        responseFailedResult.Message = uploadFailedMessage;
                        result = responseFailedResult;
                    }
                }
                else
                {
                    responseFailedResult.StatusCode = response.StatusCode;
                    responseFailedResult.Message = response.ReasonPhrase.ToString();
                    result = responseFailedResult;
                }
            }
            catch (IOException ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion 

        #region excel reader
        public object InsertExcelReader([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            string filePath = String.Empty;
            string filePathCode = String.Empty;
            string fileName = String.Empty;
            object result = null;
            string header;
            string child;
            string uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;
            var guid = Guid.NewGuid().ToString(); // bisa di ganti dengan uniq number yg lain code + tanggal
            var ctr = 0;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                //membuat folder
                var datetime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);


                    _htParameters = JsonController.setJsonToParam(jsonConvert);

                    // ini untuk dynamic spname
                    SpNames = SpName;

                    //mengambil data dari parameter yang dikirim 
                    filePathCode = guid.ToString();//_htParameters["p_file_path"].ToString();
                    fileName = _htParameters["filename"].ToString();
                    header = _htParameters["p_header"].ToString();
                    child = _htParameters["p_child"].ToString();
                    header = uploadPath + "\\" + header + "\\";
                    child = header + "\\" + child + "\\";

                    if (!Directory.Exists(header))
                    {
                        Directory.CreateDirectory(header);
                    }
                    if (!Directory.Exists(child))
                    {
                        Directory.CreateDirectory(child);
                    }

                    filePath = string.Format("{0}{1}", child, filePathCode + "_" + datetime + "_" + fileName);


                    var base64 = _htParameters["base64"].ToString();
                    byte[] byteFile = Convert.FromBase64String(base64);
                    System.IO.File.WriteAllBytes(filePath, byteFile);


                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            var headerColoumn = ds.Tables[0].Columns;
                            do
                            {
                                ctr = 0;
                                while (reader.Read()) //Each ROW
                                {

                                    if (ctr >= 1)
                                    {
                                        try
                                        {
                                            for (int j = 0; j < headerColoumn.Count; j++)
                                            {
                                                var param = headerColoumn[j];
                                                //if (!reader.GetValue(j).Equals(null))
                                                //{
                                                _htParameters["p_" + param.ToString()] = (reader.GetValue(j) ?? null);
                                                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                                                //}
                                            }
                                            _dal.Insert("", SpNames, _htParameters);

                                        }
                                        catch (Exception ex)
                                        {
                                            responseFailedResult.Data = ex.InnerException.Message.ToString();
                                            result = responseFailedResult;
                                            reader.Close();
                                            throw;
                                        }

                                    }
                                    ctr++;
                                }
                            } while (reader.NextResult()); //Move to NEXT SHEET

                        }
                    }

                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion

        #region excel reader
        public object InsertExcelReader([FromBody] dynamic data, string SpNameTable, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            Hashtable _htParameterss = null;
            string filePath = String.Empty;
            string filePathCode = String.Empty;
            string fileName = String.Empty;
            object result = null;
            string header;
            string child;
            string uploadPath = _hostingEnvironment.ContentRootPath + _docPathIfinsys;
            var guid = Guid.NewGuid().ToString(); // bisa di ganti dengan uniq number yg lain code + tanggal
            var ctr = 0;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                //membuat folder
                var datetime = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);


                    _htParameters = JsonController.setJsonToParam(jsonConvert);

                    // ini untuk dynamic spname
                    SpNames = SpName;

                    //mengambil data dari parameter yang dikirim 
                    filePathCode = guid.ToString();//_htParameters["p_file_path"].ToString();
                    fileName = _htParameters["filename"].ToString();
                    header = _htParameters["p_header"].ToString();
                    child = _htParameters["p_child"].ToString();
                    header = uploadPath + "\\" + header + "\\";
                    child = header + "\\" + child + "\\";

                    if (!Directory.Exists(header))
                    {
                        Directory.CreateDirectory(header);
                    }
                    if (!Directory.Exists(child))
                    {
                        Directory.CreateDirectory(child);
                    }

                    filePath = string.Format("{0}{1}", child, filePathCode + "_" + datetime + "_" + fileName);


                    var base64 = _htParameters["base64"].ToString();
                    byte[] byteFile = Convert.FromBase64String(base64);
                    System.IO.File.WriteAllBytes(filePath, byteFile);


                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            var headerColoumn = ds.Tables[0].Columns;
                            ctr = 0;
                            do
                            {
                                while (reader.Read()) //Each ROW
                                {

                                    if (ctr >= 1)
                                    {
                                        try
                                        {
                                            for (int j = 0; j < headerColoumn.Count; j++)
                                            {
                                                var param = headerColoumn[j];
                                                _htParameters["p_" + param.ToString()] = reader.GetValue(j);
                                                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                                            }

                                            _dal.Insert("", SpNames, _htParameters);
                                        }
                                        catch (Exception ex)
                                        {
                                            responseFailedResult.Data = ex.InnerException.Message.ToString();
                                            result = responseFailedResult;
                                            reader.Close();
                                            throw;
                                        }

                                    }
                                    ctr++;
                                }
                                //if (reader.NextResult())
                                //{

                                //}
                                dynamic parses = JArray.Parse(data.ToString());
                                _htParameterss = new Hashtable();
                                for (int j = 0; j < parse.Count; j++)
                                {
                                    var converts = JsonConvert.SerializeObject(parse[i]);

                                    string[] stringArrays = new string[] { convert };
                                    var jsonConverts = JsonConvert.SerializeObject(stringArrays);

                                    // ini untuk dynamic update bisa menerima parameter lebih dari satu dari screen
                                    _htParameterss = JsonController.setJsonToParam(jsonConvert);
                                    // ini untuk dynamic spname
                                    Utility.ApplyDefaultProp(_htParameterss, Username, IPAddress);

                                    _dal.Insert("", SpNameTable, _htParameterss);
                                }
                            } while (reader.NextResult()); //Move to NEXT SHEET
                        }
                    }

                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion

        #region DownloadFile
        public (byte[] filecontent, string contenttype, string filename) DownloadAttachment(string filename)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            try
            {
                string file_path = "";
                byte[] fileContents = null;
                string contenttype = "";

                contenttype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                file_path = _hostingEnvironment.ContentRootPath + "\\DownloadTemplate\\" + filename;

                fileContents = System.IO.File.ReadAllBytes(file_path);

                //if (declare.filecontent == null || declare.filecontent.Length == 0)
                //{

                //    return NotFound();
                //}


                return (fileContents, contenttype, filename);
            }
            catch (Exception ex)
            {
                return (null, "", ex.Message.ToString());
            }


        }
        #endregion

        #region DownloadFile with param
        public (byte[] filecontent, string contenttype, string filename) DownloadAttachmentWithParam([FromBody] dynamic data)
        {
            Hashtable _htParameters = null;

            try
            {
                string file_path = "";
                byte[] fileContents = null;
                string contenttype = "";

                _htParameters = JsonController.setJsonToParam(data.ToString());

                contenttype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                file_path = _hostingEnvironment.ContentRootPath + "\\DownloadTemplate\\" + _htParameters["p_template_name"].ToString();

                fileContents = System.IO.File.ReadAllBytes(file_path);

                //if (declare.filecontent == null || declare.filecontent.Length == 0)
                //{

                //    return NotFound();
                //}


                return (fileContents, contenttype, _htParameters["p_template_name"].ToString());
            }
            catch (Exception ex)
            {
                return (null, "", ex.Message.ToString());
            }


        }
        #endregion

        #region excel reader dynamic
        public object InsertExcelReaderDynamic([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            string filePath = String.Empty;
            string filePathCode = String.Empty;
            string fileName = String.Empty;
            object result = null;
            string uploadPath = _hostingEnvironment.ContentRootPath + "\\Files";
            var guid = Guid.NewGuid().ToString(); // bisa di ganti dengan uniq number yg lain code + tanggal
            var ctr = 0;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                //membuat folder
                //var datetime = DateTime.Now.ToString("yyyy-MM-dd");
                dynamic parse = JArray.Parse(data.ToString());
                for (int i = 0; i < parse.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(parse[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);
                    _htParameters = JsonController.setJsonToParam(jsonConvert);

                    _dt = _dal.GetRow("", SpName, _htParameters);

                    // ini untuk dynamic spname
                    // responseSuccessResult.Data["sp_upload_name"].
                    SpNames = _dt.Rows[0].ItemArray[0].ToString();

                    //mengambil data dari parameter yang dikirim
                    filePathCode = guid.ToString();//_htParameters["p_file_path"].ToString();
                    fileName = _htParameters["filename"].ToString();
                    uploadPath = uploadPath + "\\";

                    if (!Directory.Exists(uploadPath))
                    {
                        Directory.CreateDirectory(uploadPath);
                    }

                    filePath = string.Format("{0}{1}", uploadPath, filePathCode + "_" + fileName);


                    var base64 = _htParameters["base64"].ToString();
                    byte[] byteFile = Convert.FromBase64String(base64);
                    System.IO.File.WriteAllBytes(filePath, byteFile);


                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            var headerColoumn = ds.Tables[0].Columns;
                            do
                            {
                                ctr = 0;
                                while (reader.Read()) //Each ROW
                                {

                                    if (ctr >= 1)
                                    {
                                        try
                                        {
                                            //_htParameters.Clear();


                                            for (int j = 0; j < headerColoumn.Count; j++)
                                            {
                                                var param = headerColoumn[j];
                                                // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                                                _htParameters = JsonController.setJsonToParams(jsonConvert);
                                                // ini untuk dynamic spname
                                                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                                            }

                                            _dal.Insert("", SpNames, _htParameters);
                                        }
                                        catch (Exception ex)
                                        {
                                            responseFailedResult.Data = ex.InnerException.Message.ToString();
                                            result = responseFailedResult;
                                            reader.Close();
                                            throw;
                                        }

                                    }
                                    ctr++;
                                }
                            } while (reader.NextResult()); //Move to NEXT SHEET

                        }
                    }

                }
                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;
            }
            catch (Exception ex)
            {
                responseFailedResult.Data = "Row No" + "' '" + ctr.ToString() + ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }
        #endregion

        #region ExecSP
        public object ExecSpForCancelUploadDynamic([FromBody]dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            object result = null;
            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                // ini untuk dynamic insert bisa menerima parameter lebih dari satu dari screen
                _htParameters = JsonController.setJsonToParam(data.ToString());
                // ini untuk dynamic spname
                //SpName = data[0].sp_name;

                _dt = _dal.GetRow("", SpName, _htParameters);

                // ini untuk dynamic spname
                // responseSuccessResult.Data["sp_upload_name"].
                SpNames = _dt.Rows[0].ItemArray[0].ToString();

                //SpNames = SpName;
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                _dal.ExecRawSP(SpNames, _htParameters);

                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;
            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }
        #endregion

        #region mail merge
        public (byte[] filecontent, string contenttype, string filename) MailMerge([FromBody] dynamic data, string SpName)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            List<string> listNames = new List<string>();
            List<string> listValues = new List<string>();

            try
            {

                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();


                _htParameters = JsonController.setJsonToParam(data.ToString());
                //_htParameters["p_user_id"] = Username;

                FileStream fileStreamPath = new FileStream(@"DocTemplate/" + _htParameters["p_file_name"].ToString() + ".docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                _dt = _dal.GetRows("", SpName, _htParameters);

                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    string[] fieldNames = null;//{ "KOTA", "TANGGAL", "NOMOR" };
                    string[] fieldValues = null;//{ "Jakarta", "20-20-2020", "119992002", };

                    //Performs the mail 
                    foreach (DataRow row in _dt.Rows)
                    {
                        foreach (DataColumn column in _dt.Columns)
                        {
                            listNames.Add(column.ColumnName.ToString());
                            listValues.Add(row[column.ColumnName.ToString()].ToString());
                        }
                    }

                    fieldNames = listNames.ToArray();
                    fieldValues = listValues.ToArray();

                    document.MailMerge.Execute(fieldNames, fieldValues);

                    //Saves the Word document to MemoryStream
                    MemoryStream stream = new MemoryStream();
                    document.Save(stream, FormatType.Docx);
                    stream.Position = 0;
                    //Download Word document in the browser
                    return (stream.ToArray(), "application/msword", _htParameters["p_file_name"].ToString() + ".docx");
                }


            }
            catch (Exception ex)
            {
                return (null, "", ex.Message.ToString());
            }



        }
        #endregion mail merge

        #region ExecSPClient
        public object ExecSpForClientCorporateInsert([FromBody]dynamic data)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            string client_code = "";
            string bank_code = "";
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                dynamic client_main_data = JArray.Parse(data[0].client_main_data.ToString());
                dynamic client_corporate_data = JArray.Parse(data[0].client_corporate_data.ToString());
                dynamic client_corporate_notarial = JArray.Parse(data[0].client_corporate_notarial.ToString());
                dynamic client_personal_data = JArray.Parse(data[0].client_personal_data.ToString());
                dynamic client_personal_work = JArray.Parse(data[0].client_personal_work.ToString());
                dynamic client_address = JArray.Parse(data[0].client_address.ToString());
                dynamic client_asset = JArray.Parse(data[0].client_asset.ToString());
                dynamic client_bank = JArray.Parse(data[0].client_bank.ToString());
                dynamic client_bank_book = JArray.Parse(data[0].client_bank_book.ToString());
                dynamic client_doc = JArray.Parse(data[0].client_doc.ToString());
                dynamic client_sipp = JArray.Parse(data[0].client_sipp.ToString());
                dynamic client_silik = JArray.Parse(data[0].client_silik.ToString());
                dynamic client_slik_financial = JArray.Parse(data[0].client_slik_financial.ToString());
                dynamic client_financial_recapitulation = JArray.Parse(data[0].client_address.ToString());
                dynamic client_financial_statement = JArray.Parse(data[0].client_address.ToString());
                dynamic client_financial_recapitulation_detail = JArray.Parse(data[0].client_financial_recapitulation_detail.ToString());
                dynamic client_financial_statement_detail = JArray.Parse(data[0].client_financial_statement_detail.ToString());
                dynamic client_relation = JArray.Parse(data[0].client_relation.ToString());
                dynamic client_kyc = JArray.Parse(data[0].client_kyc.ToString());
                dynamic client_kyc_detail = JArray.Parse(data[0].client_kyc_detail.ToString());

                for (int i = 0; i < client_corporate_data.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_corporate_data[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Insert("", "XSP_CLIENT_CORPORATE_INFO_INSERT_FROM_CMS", _htParameters, ref client_code);
                }
                for (int i = 0; i < client_personal_data.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_personal_data[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());

                    _dal.Insert("", "XSP_CLIENT_PERSONAL_INFO_INSERT_FROM_CMS", _htParameters, ref client_code);
                }
                for (int i = 0; i < client_main_data.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_main_data[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Update("", "XSP_CLIENT_MAIN_UPDATE_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_corporate_notarial.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_corporate_notarial[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "xsp_client_corporate_notarial_insert_from_cms", _htParameters);
                }
                for (int i = 0; i < client_personal_work.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_personal_work[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_PERSONAL_WORK_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_address.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_address[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_ADDRESS_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_asset.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_asset[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_ASSET_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_bank.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_bank[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_BANK_INSERT_FROM_CMS", _htParameters, ref bank_code);

                    for (int j = 0; j < client_bank_book.Count; j++)
                    {
                        var converts = JsonConvert.SerializeObject(client_bank_book[j]);

                        string[] stringArrays = new string[] { converts };
                        var jsonConverts = JsonConvert.SerializeObject(stringArrays);

                        if (client_bank[i].code.Equals(client_bank_book[j].client_bank_code))
                        {
                            // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                            _htParameters = JsonController.setJsonToParams(jsonConverts);
                            // ini untuk dynamic spname
                            Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                            _htParameters["p_client_code"] = client_code;
                            _htParameters["p_client_bank_code"] = bank_code;

                            _dal.Insert("", "XSP_CLIENT_BANK_BOOK_INSERT_FROM_CMS", _htParameters);
                        }
                    }
                }
                for (int i = 0; i < client_doc.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_doc[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_DOC_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_sipp.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_sipp[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_SIPP_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_silik.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_silik[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_SLIK_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_slik_financial.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_slik_financial[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_SLIK_FINANCIAL_STATEMENT_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_relation.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_relation[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_RELATION_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_kyc.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_kyc[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_KYC_INSERT_FROM_CMS", _htParameters);
                }
                for (int i = 0; i < client_kyc_detail.Count; i++)
                {
                    var convert = JsonConvert.SerializeObject(client_kyc_detail[i]);

                    string[] stringArray = new string[] { convert };
                    var jsonConvert = JsonConvert.SerializeObject(stringArray);

                    // ini untuk dynamic delete bisa menerima parameter lebih dari satu dari screen
                    _htParameters = JsonController.setJsonToParams(jsonConvert);
                    // ini untuk dynamic spname
                    Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                    _htParameters["p_client_code"] = client_code;

                    _dal.Insert("", "XSP_CLIENT_KYC_DETAIL_INSERT_FROM_CMS", _htParameters);
                }

                responseSuccessResult.Data = _dal;
                responseSuccessResult.Code = client_code;
                result = responseSuccessResult;
            }
            catch (Exception ex)
            {
                DeleteClient(client_code);
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }

            return result;
        }

        public object DeleteClient(String code)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            object result = null;

            try
            {
                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                // ini untuk dynamic spname
                Utility.ApplyDefaultProp(_htParameters, Username, GetIp());
                _htParameters["p_client_code"] = code;

                _dal.Delete("", "XSP_CLIENT_MAIN_DELETE_FROM_CMS", _htParameters);

                responseSuccessResult.Data = _dal;
                result = responseSuccessResult;

            }
            catch (Exception ex)
            {
                responseFailedResult.Data = ex.InnerException.Message.ToString();
                result = responseFailedResult;
            }
            return result;
        }
        #endregion

        #region DownloadFile with data
        public (byte[] filecontent, string contenttype, string filename) DownloadExcelWithData([FromBody] dynamic data, string sp_name, string filename)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            string file_path = "";
            byte[] fileContents = null;
            string contenttype = "";

            try
            {

                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                _htParameters = JsonController.setJsonToParam(data);

                _dt = _dal.GetRows("", sp_name, _htParameters);

                XLWorkbook wb = new XLWorkbook();

                //Add DataTable in worksheet  
                wb.Worksheets.Add(_dt);

                MemoryStream stream = new MemoryStream();
                wb.SaveAs(stream);

                contenttype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


                return (stream.ToArray(), contenttype, filename);


            }
            catch (Exception ex)
            {
                return (null, "", ex.Message.ToString());
            }
        }
        public (byte[] filecontent, string contenttype, string filename) DownloadExcelWithData([FromBody] dynamic data, string sp_name)
        {
            string Username = decodeHeader(Request.Headers["UserID"]);//Request.Headers["UserID"]; // get id user
            string IPAddress = Request.Headers["IPAddress"]; // get IPAddress
            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;
            string file_path = "";
            byte[] fileContents = null;
            string contenttype = "";

            try
            {

                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                _htParameters = JsonController.setJsonToParam(data);

                _dt = _dal.GetRows("", sp_name, _htParameters);

                XLWorkbook wb = new XLWorkbook();

                //Add DataTable in worksheet  
                wb.Worksheets.Add(_dt);

                MemoryStream stream = new MemoryStream();
                wb.SaveAs(stream);

                contenttype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


                return (stream.ToArray(), contenttype, _htParameters["p_template_name"].ToString());


            }
            catch (Exception ex)
            {
                return (null, "", ex.Message.ToString());
            }
        }
        #endregion

        #region decodeHeader
        private string decodeHeader(string headerContent)
        {
            string userID = HttpUtility.UrlDecode(headerContent.ToString());//WebUtility.HtmlDecode(Request.Headers["UserID"]);
            byte[] byteUser = Convert.FromBase64String(userID);
            string decodedString = Encoding.UTF8.GetString(byteUser);
            char[] charArray = decodedString.ToCharArray();
            Array.Reverse(charArray);
            var reversedText = new String(charArray);

            return reversedText;
        }
        #endregion

        #region GetIp()
        public string GetIp()
        {
            var remoteIpAddress = Request.HttpContext.Connection.RemoteIpAddress.MapToIPv4().ToString();
            return remoteIpAddress;
        }
        #endregion GetIp()

    }
}
