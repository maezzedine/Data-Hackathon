using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class ExcelReader
    {
        ExcelPackage excelPkg;
        ExcelWorksheet sheet;

        public ExcelReader(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            excelPkg = new ExcelPackage(new FileInfo(path));
            sheet = excelPkg.Workbook.Worksheets[0];
        }

        public List<Person> GetPeople(int numberOfRows)
        {
            List<Person> People = new List<Person>();

            for (int i = 1; i < numberOfRows; i++)
            {
                var person = new Person
                {
                    Name = sheet.GetValue<String>(i, 1),
                    JobTitle = sheet.GetValue<String>(i, 17) + " - " + sheet.GetValue<String>(i, 2),
                    Location = sheet.GetValue<String>(i, 3),
                    Company = sheet.GetValue<String>(i, 4),
                    College = GetCollegeName(sheet.GetValue<String>(i, 5)),
                    LinkedinLink = sheet.GetValue<String>(i, 6),
                    Projects = sheet.GetValue<String>(i, 7),
                    Languages = sheet.GetValue<String>(i, 8),
                    JobList = sheet.GetValue<String>(i, 9)?.Split("\n").ToList(),
                    ExperienceDuration = sheet.GetValue<String>(i, 10),
                    EducationList = sheet.GetValue<String>(i, 11)?.Split("\n").ToList(),
                    CertificationList = sheet.GetValue<String>(i, 12)?.Split("\n").ToList(),
                    VolunteeringList = sheet.GetValue<String>(i, 13)?.Split("\n").ToList(),
                    VolunteeringDuration = sheet.GetValue<String>(i, 14),
                    Connections = sheet.GetValue<String>(i, 15)?.Replace("+", "").Trim(),
                    Recommendations = sheet.GetValue<String>(i, 16)
                };

                if (!person.isEmpty())
                {
                    People.Add(person);
                }
            }

            return People;
        }

        private string GetCollegeName(string college)
        {
            for (int i = 0; i < 9; i++)
            {
                var uni = (Universities)i;
                var universityName = uni.ToString().Replace("_", " ");

                if (college != null && college.Replace("-", " ").ToLower().Contains(universityName.ToLower()))
                    return universityName;
            }

            return null;
        }
    }
}
