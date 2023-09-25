using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Threading.Tasks;
using TrainingAPI.App_Code;

namespace TrainingAPI.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class MasterProjectPositionController : iFinancingController
    {
        public MasterProjectPositionController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : base(configuration, hostingEnvironment)
        {

        }

        private string _tableName = "MASTER_PROJECT_POSITION";
        private string _spName = "";

        [HttpPost]
        [Route("GetRows")]
        public async Task<IActionResult> GetRows([FromBody] dynamic data)
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
        public async Task<IActionResult> GetRow([FromBody] dynamic data)
        {
            return await Task.Run(() =>
            {
                return Ok(base.GetRow((JArray)data, _tableName));
            });
        }

        [HttpPost]
        [Route("Insert")]
        public async Task<IActionResult> Insert([FromBody] dynamic data)
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
    }
}