using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using RegistroDeCompetenciaBash.Data;
using RegistroDeCompetenciaBash.Models;

namespace RegistroDeCompetenciaBash
{
    class Program
    {
        static int tableRow = 1;
        static async Task Main(string[] args)
        {
            CreateExcel(await DbContext.instance.SPGetStudents());
        }

        // ------------------------------------ Methods --------------------------------------- //
        static public void CreateExcel(IEnumerable<Estudiante> estudiantes)
        {
            string fileName = "RegistroDeCompetencia";

            using (ExcelPackage excel = new ExcelPackage())
            {
                try
                {
                    excel.Workbook.Worksheets.Add(fileName);
                    GetTableHeaders(excel.Workbook.Worksheets[fileName]);
                    PopulateTable(excel.Workbook.Worksheets[fileName], estudiantes);

                    FileInfo excelFile = new FileInfo(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx" );

                    excel.SaveAs(excelFile);
                }
                catch(UnauthorizedAccessException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static public void GetTableHeaders(ExcelWorksheet worksheet)
        {
            int count = 1;

            foreach(var item in typeof(Estudiante).GetProperties())
            {
                if(item.Name.ToUpper() == "ID")
                {
                    worksheet.Cells[1, count].Value = "Numero De Estudiante";
                    count++;
                    continue;
                }
                else if(item.Name.ToUpper() == "RECINTOID")
                {
                    continue;
                }
                else
                {
                    worksheet.Cells[1, count].Value = item.Name;
                }
                count++;
            }
            
            tableRow++;
        }

        static public void PopulateTable(ExcelWorksheet worksheet, IEnumerable<Estudiante> estudiantes)
        {
            foreach(var estudiante in estudiantes)
            {
                worksheet.Cells[tableRow, 1].Value = estudiante.Id;
                worksheet.Cells[tableRow, 2].Value = estudiante.Nombre;
                worksheet.Cells[tableRow, 3].Value = estudiante.ApellidoPaterno;
                worksheet.Cells[tableRow, 4].Value = estudiante.ApellidoMaterno;
                worksheet.Cells[tableRow, 5].Value = estudiante.Email;
                worksheet.Cells[tableRow, 6].Value = estudiante.Recinto.Nombre;
                tableRow++;
            }
        }

    }
}
