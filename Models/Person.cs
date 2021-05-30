using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Models
{
    public class Person
    {
        public string Name { get; set; }
        public string JobTitle { get; set; }
        public string Location { get; set; }
        public string Company { get; set; }
        public string College { get; set; }
        public string LinkedinLink { get; set; }
        public string Projects { get; set; }
        public string Languages { get; set; }
        public string Connections { get; set; }
        public string Recommendations { get; set; }
        public List<string> JobList { get; set; }
        public string Jobs
        {
            get
            {
                return JobList?.Aggregate("", (prev, current) => prev + " \n" + current)?.Replace("\r\n", "");
            }
        }
        public string ExperienceDuration { get; set; }
        public List<string> EducationList { get; set; }
        public string Education
        {
            get
            {
                return EducationList?.Aggregate("", (prev, current) => prev + " \n" + current)?.Replace("\r\n", "");
            }
        }
        public List<string> CertificationList { get; set; }
        public string Certification
        {
            get
            {
                return CertificationList?.Aggregate("", (prev, current) => prev + " \n" + current)?.Replace("\r\n", "");
            }
        }
        public List<string> VolunteeringList { get; set; }
        public string Volunteering
        {
            get
            {
                return VolunteeringList?.Aggregate("", (prev, current) => prev + " \n" + current)?.Replace("\r\n", "");
            }
        }
        public string VolunteeringDuration { get; set; }
        public List<string> SkillsList { get; set; }
        public string Skills
        {
            get
            {
                return SkillsList.Aggregate("", (prev, current) => prev + " \n" + current).Replace("\r\n", "");
            }
        }

        public void WriteToExcel(ExcelWorksheet sheet, int row)
        {
            sheet.Cells[row, 1].Value = Name;
            sheet.Cells[row, 2].Value = JobTitle;
            sheet.Cells[row, 3].Value = Location;
            sheet.Cells[row, 4].Value = Company;
            sheet.Cells[row, 5].Value = College;
            sheet.Cells[row, 6].Value = LinkedinLink;
            sheet.Cells[row, 7].Value = Projects;
            sheet.Cells[row, 8].Value = Languages;
            sheet.Cells[row, 9].Value = Jobs;
            sheet.Cells[row, 10].Value = ExperienceDuration;
            sheet.Cells[row, 11].Value = Education;
            sheet.Cells[row, 12].Value = Certification;
            sheet.Cells[row, 13].Value = Volunteering;
            sheet.Cells[row, 14].Value = VolunteeringDuration;
            sheet.Cells[row, 15].Value = Connections;
            sheet.Cells[row, 16].Value = Recommendations;
        }

        public bool isEmpty()
        {
            return string.IsNullOrEmpty(Name);
        }

        public IEnumerable<string> GetCompanies()
        {
            if (JobList == null)
                return new List<string>();

            try
            {
                return JobList.Where(j => j.Split(" - ").Length > 2).Select(j => j.Split(" - ")[1]);
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }

        public IEnumerable<string> GetNgos()
        {
            if (VolunteeringList == null)
                return new List<string>();

            try
            {
                return VolunteeringList.Where(v => v.Split(" - ").Length > 2).Select(v => v.Split(" - ")[1]);
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }
    }
}
