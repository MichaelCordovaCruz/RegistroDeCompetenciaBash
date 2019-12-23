using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;

using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;

using RegistroDeCompetenciaBash.Models;

namespace RegistroDeCompetenciaBash
{
    class Program
    {
        static async Task Main(string[] args)
        {
            List<Estudiante> estudiantes = await GetStudentDataAsync();
            CreateExcel(estudiantes);
        }

        static public async Task<List<Estudiante>> GetStudentDataAsync()
        {
            var response = await new HttpClient().GetStringAsync("https://inter-venture-competition-2020.herokuapp.com/api");
            return JsonConvert.DeserializeObject<List<Estudiante>>(response);
        }

        static public void CreateExcel(List<Estudiante> estudiantes)
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
        }

        static public void PopulateTable(ExcelWorksheet worksheet, List<Estudiante> estudiantes)
        {
            int count = 2;

            foreach(var estudiante in estudiantes)
            {
                worksheet.Cells[count, 1].Value = estudiante.Id;
                worksheet.Cells[count, 2].Value = estudiante.Nombre;
                worksheet.Cells[count, 3].Value = estudiante.ApellidoPaterno;
                worksheet.Cells[count, 4].Value = estudiante.ApellidoMaterno;
                worksheet.Cells[count, 5].Value = estudiante.Email;
                worksheet.Cells[count, 6].Value = estudiante.Recinto.Nombre;
                count++;
            }
        }

    }
}
