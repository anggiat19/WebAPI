using DataAccessLayer;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using TrainingAPI.Interfaces;

namespace TrainingAPI.App_Code
{
    public class UserManagementService : IUserManagementService
    {
        private readonly string _connectionString;

        public UserManagementService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        public bool IsValidUser(string userName, string password)
        {

            GeneralDAL _dal = null;
            Hashtable _htParameters = null;
            DataTable _dt = null;


            try
            {

                _dal = new GeneralDAL(_connectionString);
                _htParameters = new Hashtable();

                _htParameters["p_uid"] = userName.ToString();
                _htParameters["p_password"] = password.ToString();

                _dt = _dal.GetRow("SYS_USER_MAIN", "xsp_master_user_main_validate", _htParameters);

                if (_dt.Rows.Count == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }

                //return Ok(JsonConvert.SerializeObject(_dt, Formatting.Indented));

            }
            catch (Exception ex)
            {
                //return BadRequest(new { status = 0, message = ex.InnerException.Message });
                return false;
            }

            //return true;

            
        }
    }
}
