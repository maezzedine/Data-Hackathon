using Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ScrapingData
{
    public class ExcelCleaner : IDisposable
    {
        ExcelPackage excelPkg;
        ExcelWorksheet sheet;
        List<Person> People = new List<Person>();
        string outputPath;

        public ExcelCleaner(string path)
        {
            outputPath = $"Cleaned {path}";

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            excelPkg = new ExcelPackage(new FileInfo(path));
            sheet = excelPkg.Workbook.Worksheets[0];
            sheet.Cells[1, 1].Value = "Name";
            sheet.Cells[1, 2].Value = "Job";
            sheet.Cells[1, 3].Value = "Location";
            sheet.Cells[1, 4].Value = "Company";
            sheet.Cells[1, 5].Value = "College";
            sheet.Cells[1, 6].Value = "LinkedinLink";
            sheet.Cells[1, 7].Value = "Projects";
            sheet.Cells[1, 8].Value = "Languages";
            sheet.Cells[1, 9].Value = "Experience";
            sheet.Cells[1, 10].Value = "ExperienceDuration";
            sheet.Cells[1, 11].Value = "Education";
            sheet.Cells[1, 12].Value = "Certification";
            sheet.Cells[1, 13].Value = "Volunteering";
            sheet.Cells[1, 14].Value = "VolunteeringDuration";
            sheet.Cells[1, 15].Value = "Connections";
            sheet.Cells[1, 16].Value = "Recommendations";
        }

        public void Clean()
        {
            for (int i = 2; i < 1566; i++)
            {
                try
                {
                    var experienceDuration = new TimeDuration(sheet.GetValue<String>(i, 10));
                    var volunteeringDuration = new TimeDuration(sheet.GetValue<String>(i, 14));

                    var person = new Person
                    {
                        Name = sheet.GetValue<String>(i, 1),
                        JobTitle = sheet.GetValue<String>(i, 2),
                        Location = sheet.GetValue<String>(i, 3),
                        Company = sheet.GetValue<String>(i, 4),
                        College = sheet.GetValue<String>(i, 5),
                        LinkedinLink = sheet.GetValue<String>(i, 6),
                        Projects = sheet.GetValue<String>(i, 7),
                        Languages = sheet.GetValue<String>(i, 8),
                        JobList = sheet.GetValue<String>(i, 9)?.Split("\n").ToList(),
                        ExperienceDuration = experienceDuration.getMonths().ToString(),
                        EducationList = sheet.GetValue<String>(i, 11)?.Split("\n").ToList(),
                        CertificationList = sheet.GetValue<String>(i, 12)?.Split("\n").ToList(),
                        VolunteeringList = sheet.GetValue<String>(i, 13)?.Split("\n").ToList(),
                        VolunteeringDuration = volunteeringDuration.getMonths().ToString(),
                        Connections = sheet.GetValue<String>(i, 15)?.Replace("connection", "").Trim(),
                        Recommendations = sheet.GetValue<String>(i, 16)
                    };

                    if (!person.isEmpty())
                    {
                        People.Add(person);
                    }
                }
                catch (Exception e) 
                { 
                }
            }

            sheet.Dispose();
        }

        public void Dispose()
        {
            excelPkg.Save();
            excelPkg.Dispose();
        }

        public void Save()
        {
            excelPkg = new ExcelPackage(new FileInfo(outputPath));
            sheet = excelPkg.Workbook.Worksheets.Add("Sheet 1");

            sheet.Cells[1, 1].Value = "Name";
            sheet.Cells[1, 2].Value = "Job";
            sheet.Cells[1, 3].Value = "Location";
            sheet.Cells[1, 4].Value = "Company";
            sheet.Cells[1, 5].Value = "College";
            sheet.Cells[1, 6].Value = "LinkedinLink";
            sheet.Cells[1, 7].Value = "Projects";
            sheet.Cells[1, 8].Value = "Languages";
            sheet.Cells[1, 9].Value = "Experience";
            sheet.Cells[1, 10].Value = "ExperienceDuration";
            sheet.Cells[1, 11].Value = "Education";
            sheet.Cells[1, 12].Value = "Certification";
            sheet.Cells[1, 13].Value = "Volunteering";
            sheet.Cells[1, 14].Value = "VolunteeringDuration";
            sheet.Cells[1, 15].Value = "Connections";
            sheet.Cells[1, 16].Value = "Recommendations";

            for (int i = 0; i < People.Count; i++)
            {
                var person = People[i];
                person.WriteToExcel(sheet, i + 2);
            }
        }
    }
}
