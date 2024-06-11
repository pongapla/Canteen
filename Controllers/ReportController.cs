using ClosedXML.Excel;
using iText.Kernel.Pdf;
using Microsoft.AspNetCore.Mvc;
using System.Data.OleDb;
using TLM_Canteen.Models;
using iText.Layout.Element;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Layout.Properties;
using iText.Layout;



namespace TLM_Canteen.Controllers;

public class ReportController : Controller
{
    public IActionResult Index(string? startDate, string? endDate)
    {
        IConfiguration configuration = new ConfigurationBuilder()
               .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
               .Build();
        string connectionString = configuration.GetConnectionString("AccessConnection");
        List<Report> ReportList = new List<Report>();
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {

            try
            {
                string query = @"SELECT DISTINCT 
                        FORMAT(ti.[DateTime], 'dd-MM-yyyy') AS _DateTime,
                        COUNT(ti.CodeEm) AS TotalScan,
                        COUNT(ti.CodeEm) * (SELECT Price FROM [TLM_Price] WHERE CodeP = '001') AS TotalPrice
                        FROM 
                            [TLM_TimeIn] ti
                        WHERE ti.[DateTime] BETWEEN ? AND DATEADD(""d"", 1,?)
                        GROUP BY 
                            FORMAT(ti.[DateTime], 'dd-MM-yyyy')";

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
                            var report = new Report()
                            {
                                
                                _Date = reader["_DateTime"].ToString(),
                                TotalScan = reader["TotalScan"].ToString(),
                                TotalPrice = reader["TotalPrice"].ToString(),
                            };
                            ReportList.Add(report);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exception (e.g., log the error)
                ex.Message.ToString();

            }
        }

        return View(ReportList);
    }

    public async Task<IActionResult> ExportExcel([FromBody] List<Report> modelData)
    {
        try
        {
            if (modelData == null || modelData.Count == 0)
            {
                return BadRequest("No data provided.");
            }

            var firstDate = modelData.First()._Date.ToString();
            var lastDate = modelData.Last()._Date.ToString();

            // Calculate total scan and total price
            var totalScan = modelData.Sum(d => int.Parse(d.TotalScan));
            var totalPrice = modelData.Sum(d => int.Parse(d.TotalPrice));


            // สร้างไฟล์ Excel
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");
            // สร้างแถวหัวตาราง
            worksheet.Range("A1:F1").Merge().Value = $"ข้อมูลการสแกนนิ้วตั้งแต่วันที่ {firstDate} ถึง {lastDate}";
            worksheet.Range("A1:F1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Set the header row (Row 2 after merging)
            worksheet.Cell(2, 1).Value = "วันที่";
            worksheet.Cell(2, 2).Value = "จำนวนที่สแกนนิ้ว";
            worksheet.Cell(2, 3).Value = "รวมเป็นเงิน";
            worksheet.Cell(2, 4).Value = "ลายเซนต์ HR";
            worksheet.Cell(2, 5).Value = "ลายเซนต์แม่ค้า";
            worksheet.Cell(2, 6).Value = "หมายเหตุ";

            // Insert the data starting from Row 3
            int row = 3;
            foreach (var item in modelData)
            {
                worksheet.Cell(row, 1).Value = item._Date.ToString();
                worksheet.Cell(row, 2).Value = item.TotalScan;
                worksheet.Cell(row, 3).Value = item.TotalPrice;
                worksheet.Cell(row, 4).Value = "";
                worksheet.Cell(row, 5).Value = "";
                worksheet.Cell(row, 6).Value = "";
                row++;
            }

            // Add total row
            worksheet.Cell(row, 1).Value = "Total:";
            worksheet.Cell(row, 2).Value = totalScan;
            worksheet.Cell(row, 3).Value = totalPrice.ToString("N0");
            worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
            worksheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0";

           
            // บันทึกไฟล์ Excel ลง MemoryStream
            var stream = new MemoryStream();

            workbook.SaveAs(stream);
            stream.Position = 0;

            // ส่งไฟล์ Excel กลับไปยังไคลเอ็นต์
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SumCanteen.xlsx");
        }
        catch (Exception ex)
        {
            return StatusCode(500, "An error occurred while processing the request.");
        }
    }

    [HttpPost]
    public async Task<IActionResult> ExportPDF([FromBody] List<Report> modelData)
    {
        if (modelData == null || modelData.Count == 0)
        {
            return BadRequest("No data provided.");
        }

        try
        {
            // Create PdfFont from THSarabunNew.ttf font
            string fontPath = @"D:\General\DataBaseCanteen\TH Sarabun\TH Sarabun PSK V-1\THSarabun.ttf";
            PdfFont font = PdfFontFactory.CreateFont(fontPath, PdfEncodings.IDENTITY_H);

            // สร้างตัวแปรสำหรับรูปแบบข้อความในเฮดเดอร์
            Style headerStyle = new Style()
                .SetFont(font)
                .SetFontSize(12)
                .SetBold();

            using (var stream = new MemoryStream())
            {
                // Create PdfWriter and PdfDocument
                PdfWriter writer = new PdfWriter(stream);
                PdfDocument pdf = new PdfDocument(writer);

                // Create iText.Layout.Document
                Document document = new Document(pdf);

                var firstDate = modelData.First()._Date.ToString();
                var lastDate = modelData.Last()._Date.ToString();
                var totalScan = modelData.Sum(d => int.Parse(d.TotalScan));
                var totalPrice = modelData.Sum(d => int.Parse(d.TotalPrice));

                // Title
                Paragraph title = new Paragraph($"ข้อมูลการสแกนนิ้วตั้งแต่วันที่ {firstDate} ถึง {lastDate}")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20)
                    .SetFont(font);
                document.Add(title);

                // Table
                Table table = new Table(new float[] { 3, 3, 3 });
                table.SetWidth(UnitValue.CreatePercentValue(100));

                // Table Header
                table.AddHeaderCell(new Cell().Add(new Paragraph("วันที่").AddStyle(headerStyle)));
                table.AddHeaderCell(new Cell().Add(new Paragraph("จำนวนที่สแกนนิ้ว").AddStyle(headerStyle)));
                table.AddHeaderCell(new Cell().Add(new Paragraph("รวมเป็นเงิน").AddStyle(headerStyle)));

                // Table Content
                foreach (var item in modelData)
                {
                    table.AddCell(new Cell().Add(new Paragraph(item._Date).SetFont(font)));
                    table.AddCell(new Cell().Add(new Paragraph(item.TotalScan).SetFont(font)));
                    table.AddCell(new Cell().Add(new Paragraph(item.TotalPrice).SetFont(font)));
                }

                // Add total row
                table.AddCell(new Cell().Add(new Paragraph("Total:").SetFont(font)));
                table.AddCell(new Cell().Add(new Paragraph(totalScan.ToString("N0")).SetFont(font))); // formatted with thousands separator
                table.AddCell(new Cell().Add(new Paragraph(totalPrice.ToString("N0")).SetFont(font))); // formatted with thousands separator

                // Add table to document
                document.Add(table);

                // Signature
                Paragraph signature = new Paragraph("คุณเดชาวัต โพธิ์วงค์\n(ผู้จัดการแผนกทรัพยากรมนุษย์และการบริหาร)")
                    .SetTextAlignment(TextAlignment.CENTER) // จัดวางให้ตรงกลางทั้งแนวตั้งและแนวนอน
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetMarginTop(40)
                    .SetFont(font);
                document.Add(signature);

                document.Close();

                return File(stream.ToArray(), "application/pdf", "SumCanteen.pdf");
            }

        }
        catch (Exception ex)
        {
            // Handle exception
            return BadRequest($"Error: {ex.Message}");
        }
    }
}

