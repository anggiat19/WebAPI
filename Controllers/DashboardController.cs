using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using TrainingAPI.App_Code;
using Microsoft.AspNetCore.Authorization;

namespace TrainingAPI.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class DashboardController : iFinancingController
    {
        public DashboardController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : base(configuration, hostingEnvironment)
        {

        }

        private string _tableName = "CALENDER_EVENT";
        private string _spNameForEvent = "XSP_CALENDER_EVENT_GETROW";
        private string _spNameForGetRole = "";
        private string _spName = "";

        [HttpPost]
        [Route("GetRows")]
        public async Task<IActionResult> GetRows([FromBody]dynamic data)
        {
            return await Task.Run(() =>
            {
                var Jsondata = JsonConvert.SerializeObject(data);
                JObject JsonObject = JObject.Parse(Jsondata);

                return Ok(base.GetRows(JsonObject, _tableName));
            });
        }

        [HttpPost]
        [Route("GetRow")]
        public async Task<IActionResult> GetRow([FromBody]dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.GetRow((JArray)data, _tableName));
            });
        }

        [HttpPost]
        [Route("Insert")]
        public async Task<IActionResult> Insert([FromBody]dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.Insert((JArray)data, _tableName));
            });
        }

        [HttpPost]
        [Route("Update")]
        public async Task<IActionResult> Update([FromBody] dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.Update((JArray)data, _tableName));
            });
        }

        [HttpPost]
        [Route("Delete")]
        public async Task<IActionResult> Delete([FromBody] dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.Delete((JArray)data, _tableName));
            });
        }

        [HttpPost]
        [Route("ExecSpForEvent")]
        public async Task<IActionResult> ExecSp([FromBody]dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.ExecSp((JArray)data, _spNameForEvent));
            });
        }


        [HttpPost]
        [Route("ExecSpForGetRole")]
        public async Task<IActionResult> ExecSpForGetRole([FromBody]dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.ExecSp((JArray)data, _spNameForGetRole));
            });
        }
    }
}
