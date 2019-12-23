

namespace RegistroDeCompetenciaBash
{
    public class GenerateExcel
    {
        // public IActionResult Download()
        // {
        //     byte[] fileContents;

        //     using (var package = new ExcelPackage())
        //     {
        //         var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        //         // Put whatever you want here in the sheet
        //         // For example, for cell on row1 col1
        //         worksheet.Cells[1, 1].Value = "Long text";

        //         worksheet.Cells[1, 1].Style.Font.Size = 12;
        //         worksheet.Cells[1, 1].Style.Font.Bold = true;

        //         worksheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

        //         // So many things you can try but you got the idea.

        //         // Finally when you're done, export it to byte array.
        //         fileContents = package.GetAsByteArray();
        //     }

        //     if (fileContents == null || fileContents.Length == 0)
        //     {
        //         return NotFound();
        //     }

        //     return File(
        //         fileContents: fileContents,
        //         contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        //         fileDownloadName: "test.xlsx"
        //     );
        // }

        // using (ExcelPackage excel = new ExcelPackage())
        // {
        //     excel.Workbook.Worksheets.Add("Worksheet1");
        //     excel.Workbook.Worksheets.Add("Worksheet2");
        //     excel.Workbook.Worksheets.Add("Worksheet3");
            
        //     var headerRow = new List<string[]>()
        //     {
        //         new string[] { "ID", "First Name", "Last Name", "DOB" }
        //     };
            
        //     // Determine the header range (e.g. A1:D1)
        //     string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

        //     // Target a worksheet
        //     var worksheet = excel.Workbook.Worksheets["Worksheet1"];
            
        //     // Popular header row data
        //     worksheet.Cells[headerRange].LoadFromArrays(headerRow);
            
        //     FileInfo excelFile = new FileInfo(@"C:\Users\amir\Desktop\test.xlsx");
        //     excel.SaveAs(excelFile);
        // }
    }
}