using Microsoft.AspNetCore.Mvc;
using System.Data.OleDb;
using TLM_Canteen.Models;
using TLM_Canteen.Services;
using ClosedXML.Excel;




namespace TLM_Canteen.Controllers
{
    public class HomeController : Controller
    {
        private readonly AccessDbService _dbService;


        public HomeController(AccessDbService dbService)
        {
            _dbService = dbService;
        }

        public IActionResult Index(string startDate, string endDate)
        {
            IConfiguration configuration = new ConfigurationBuilder()
               .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
               .Build();

            string connectionString = configuration.GetConnectionString("AccessConnection");
            List<SearchUser> userList = new List<SearchUser>();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {

                try
                {
                    string query = @"SELECT
                                        te.[EmpCode] AS Code,
                                        'คุณ ' & te.[EmpFNameThai] & ' ' & te.[EmpLNameThai] AS Name,
                                        te.[OrgTDesc] AS Department,
                                        Format(ti.[DateTime], ""dd-MM-yyyy"") AS _DateTime,
                                        ti.[Time] AS _Time,
                                        COUNT(*) AS Total
                                    FROM
                                        [TLM_TimeIn] ti
                                    LEFT JOIN
                                        [TLM_Employee] te ON ti.[CodeEm] = te.[EmpCode]
                                    WHERE
                                        ti.[DateTime] BETWEEN ? AND DATEADD(""d"", 1,?)
                                    GROUP BY
                                        te.[EmpCode],
                                        te.[EmpFNameThai],
                                        te.[EmpLNameThai],
                                        te.[OrgTDesc],
                                        Format(ti.[DateTime], ""dd-MM-yyyy""),
                                        ti.[Time]
                                    ";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.Add(new OleDbParameter(" ? ", OleDbType.Date) { Value = DateTime.Parse(startDate) });
                        command.Parameters.Add(new OleDbParameter("?", OleDbType.Date) { Value = DateTime.Parse(endDate) });

                        connection.Open();
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            int i = 1;
                            while (reader.Read())
                            {
                                var user = new SearchUser()
                                {
                                    No = i++,
                                    Code = reader["Code"].ToString(),
                                    Name = reader["Name"].ToString(),
                                    Department = reader["Department"].ToString(),
                                    _DateTime = reader["_DateTime"].ToString(),
                                    _Time = reader["_Time"].ToString(),
                                    Total = reader["Total"] != DBNull.Value ? (int?)Convert.ToInt32(reader["Total"]) : null
                                };
                                userList.Add(user);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exception (e.g., log the error)
                }
            }

            return View(userList);
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
                String _DateStart;
                DateTime _DateEnd;
                try
                {


                    // ดึงค่า _DateStart จากฐานข้อมูล
                    string query = @"SELECT TOP 1 Datetime FROM TLM_TimeIn Order By DateTime DESC";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        connection.Open();
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read()) // ตรวจสอบว่ามีข้อมูลหรือไม่
                            {
                                _DateStart = reader[0].ToString();
                            }
                            else
                            {
                                throw new Exception("No data found in TLM_TimeIn table.");
                            }
                        }
                        connection.Close();
                    }

                    // กำหนดค่า _DateEnd ให้เป็นเวลาปัจจุบัน
                    _DateEnd = DateTime.Now;
                    string FormateDate = _DateEnd.ToString("yyyy-MM-dd");
                    // ดึงข้อมูลจาก CHECKINOUT ที่อยู่ในช่วงเวลา _DateStart ถึง _DateEnd
                    string query2 = @"SELECT Badgenumber, CHECKTIME FROM CHECKINOUT WHERE CHECKTIME BETWEEN ? AND ?";

                    using (OleDbCommand command = new OleDbCommand(query2, connection))
                    {
                        connection.Open();
                        // กำหนดค่าพารามิเตอร์
                        command.Parameters.Add(new OleDbParameter("?", _DateStart));
                        command.Parameters.Add(new OleDbParameter("?", FormateDate));

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {

                            while (reader.Read())
                            {
                                // ดำเนินการกับข้อมูลที่ดึงมา
                                var check = new CheckInout()
                                {
                                    Code = reader["Badgenumber"].ToString(),
                                    _DateTime = Convert.ToDateTime(reader["CHECKTIME"])
                                };
                                checkList.Add(check);
                            }

                            // ทำสิ่งที่คุณต้องการกับ checkList
                        }
                        connection.Close();
                    }

                    string query3 = @"INSERT INTO TLM_TimeIn ([DateTime], CodeEm, [Time])
                            VALUES (?,?,?)";

                    foreach (var check in checkList)
                    {

                        connection.Open();
                        using (OleDbCommand command = new OleDbCommand(query3, connection))
                        {
                            string checkDuplicate = @"SELECT  count([CodeEm]) FROM [TLM_TimeIn] WHERE [CodeEm] = ? AND [DateTime] = ?";
                            int rowCount = 0;
                            using (OleDbCommand checkCommand = new OleDbCommand(checkDuplicate, connection))
                            {
                                checkCommand.Parameters.Add(new OleDbParameter("CodeEm", OleDbType.VarChar)).Value = check.Code;
                                checkCommand.Parameters.Add(new OleDbParameter("DateTime", OleDbType.Date)).Value = check._DateTime;
                                rowCount = (int)checkCommand.ExecuteScalar();
                            }

                            if (check.Code.Length < 4)
                            {
                                check.Code = check.Code.PadLeft(4, '0');
                            }
                            
                            command.Parameters.Add(new OleDbParameter("DateTime", OleDbType.Date)).Value = check._DateTime; // yyyy, MM, dd, HH, mm, ss
                            command.Parameters.Add(new OleDbParameter("CodeEm", OleDbType.VarChar)).Value = check.Code;
                            command.Parameters.Add(new OleDbParameter("Time", OleDbType.VarChar)).Value = check._DateTime.TimeOfDay.ToString();


                            if (rowCount > 0)
                            {
                                Console.WriteLine("Duplicate data found. No data inserted.");
                            }
                            else
                            {
                                int rowsAffected = command.ExecuteNonQuery();
                            }

                        }
                        connection.Close();
                    }
                    // ตรวจสอบข้อมูลซ้ำและเปลี่ยนสถานะ เป็น Inactive

                    string query4 = @"INSERT INTO TLM_Report ([DateTime], CodeEm, [Time])
                            VALUES (?,?,?)";
                    foreach (var check in checkList)
                    {
                        connection.Open();
                        using (OleDbCommand command = new OleDbCommand(query4, connection))
                        {
                            string checkDuplicate2 = "SELECT count([CodeEm]) FROM TLM_Report WHERE [CodeEm] = ? AND DateValue([DateTime]) = ?";
                            int rowCount2 = 0;
                            using (OleDbCommand checkCommand = new OleDbCommand(checkDuplicate2, connection))
                            {
                                var _NewDate = check._DateTime.ToString("yyyy-MM-dd");
                                checkCommand.Parameters.Add(new OleDbParameter("CodeEm", OleDbType.VarChar)).Value = check.Code;
                                checkCommand.Parameters.Add(new OleDbParameter("DateTime", OleDbType.Date)).Value = _NewDate.ToString();
                                rowCount2 = (int)checkCommand.ExecuteScalar();
                            }

                            command.Parameters.Add(new OleDbParameter("DateTime", OleDbType.Date)).Value = check._DateTime; // yyyy, MM, dd, HH, mm, ss
                            command.Parameters.Add(new OleDbParameter("CodeEm", OleDbType.VarChar)).Value = check.Code;
                            command.Parameters.Add(new OleDbParameter("Time", OleDbType.VarChar)).Value = check._DateTime.TimeOfDay.ToString();

                            if (rowCount2 > 0)
                            {
                                Console.WriteLine("Duplicate data found. No data updated.");
                                
                            }
                            else
                            {
                                int rowsAffected = command.ExecuteNonQuery();
                            }

                        }
                        connection.Close();
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

        public IActionResult SearchUser(string startDate, string endDate)
        {
            return RedirectToAction("Index", "Home", new { startDate, endDate });
        }

        public async Task<IActionResult> ExportExcel([FromBody] List<SearchUser> modelData)
        {
            try
            {
                if (modelData == null || modelData.Count == 0)
                {
                    return BadRequest("No data provided.");
                }

                // สร้างไฟล์ Excel
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Data");
                // สร้างแถวหัวตาราง

                worksheet.Cell(1, 1).Value = "No";
                worksheet.Cell(1, 2).Value = "Code";
                worksheet.Cell(1, 3).Value = "Name";
                worksheet.Cell(1, 4).Value = "Department";
                worksheet.Cell(1, 5).Value = "_DateTime";
                worksheet.Cell(1, 6).Value = "_Time";
                worksheet.Cell(1, 7).Value = "Total";

                int row = 2;
                foreach (var item in modelData)
                {
                    worksheet.Cell(row, 1).Value = item.No;
                    worksheet.Cell(row, 2).Value = item.Code;
                    worksheet.Cell(row, 3).Value = item.Name;
                    worksheet.Cell(row, 4).Value = item.Department;
                    worksheet.Cell(row, 5).Value = item._DateTime;
                    worksheet.Cell(row, 6).Value = item._Time;
                    worksheet.Cell(row, 7).Value = item.Total;
                    row++;
                }
                // บันทึกไฟล์ Excel ลง MemoryStream
                var stream = new MemoryStream();

                workbook.SaveAs(stream);
                stream.Position = 0;

                // ส่งไฟล์ Excel กลับไปยังไคลเอ็นต์
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "data.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, "An error occurred while processing the request.");
            }
        }

    }
}