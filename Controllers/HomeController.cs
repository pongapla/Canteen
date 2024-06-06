using Microsoft.AspNetCore.Mvc;
using System.Data.OleDb;
using System.Diagnostics;
using TLM_Canteen.Models;
using TLM_Canteen.Services;

namespace TLM_Canteen.Controllers
{
    public class HomeController : Controller
    {
        private readonly AccessDbService _dbService;

        public HomeController(AccessDbService dbService)
        {
            _dbService = dbService;
        }

        public IActionResult Index()
        {
            // เรียกใช้บริการเพื่อดึงข้อมูล

            return View();
        }

        public IActionResult GetDataInSert()
        {
            IConfiguration configuration = new ConfigurationBuilder()
              .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
              .Build();

            string connectionString = configuration.GetConnectionString("AccessConnection");

            List<CheckInout> checkList = new List<CheckInout>();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string _DateStart;
                DateTime _DateEnd;
                try
                {
                    connection.Open();

                    // ดึงค่า _DateStart จากฐานข้อมูล
                    string query = @"SELECT TOP 1 Datetime FROM TLM_TimeIn Order By DateTime DESC";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read()) // ตรวจสอบว่ามีข้อมูลหรือไม่
                            {
                                _DateStart = reader["Datetime"].ToString(); // ดึงค่า DateTime จากคอลัมน์ที่ 0
                            }
                            else
                            {
                                throw new Exception("No data found in TLM_TimeIn table.");
                            }
                        }
                    }

                    // กำหนดค่า _DateEnd ให้เป็นเวลาปัจจุบัน
                    _DateEnd = DateTime.Now;
                    string FormateDate = _DateEnd.ToString();
                    // ดึงข้อมูลจาก CHECKINOUT ที่อยู่ในช่วงเวลา _DateStart ถึง _DateEnd
                    string query2 = @"SELECT Badgennumber, CHECKTIME FROM CHECKINOUT WHERE CHECKTIME BETWEEN ? AND ?;";

                    using (OleDbCommand command = new OleDbCommand(query2, connection))
                    {
                        command.Parameters.AddWithValue("?", _DateStart);
                        command.Parameters.AddWithValue("?", FormateDate);

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {

                            while (reader.Read())
                            {
                                // ดำเนินการกับข้อมูลที่ดึงมา
                                var check = new CheckInout()
                                {
                                    Code = reader["Badgennumber"].ToString(),
                                    _DateTime = Convert.ToDateTime(reader["CHECKTIME"])
                                };
                                checkList.Add(check);
                            }

                            // ทำสิ่งที่คุณต้องการกับ checkList
                        }
                    }
                }
                catch (Exception ex)
                {
                    // จัดการข้อผิดพลาดที่เกิดขึ้น
                    Console.WriteLine("An error occurred: " + ex.Message);
                }

                return Ok("Ok");
            }

        }
    }
}
