using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using HtmlAgilityPack;
using Models;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace ScrapingData
{
    class Program
    {
        static ChromeDriver driver;
        static string ChromDriverDirectory = @"E:\Applications\Setup\chromedriver_win32";
        static string OutputFileName = "data.xls";
        static ExcelPackage excelPkg;
        static ExcelWorksheet sheet;
        static List<Person> People = new List<Person>();
        static string sheetName = "Sheet57";
        static int graduationYearThreshold = 2015;
        static string email = Constants.Accounts.TwentySeven;
        static string password = Constants.Accounts.Password_2;
        static int offset = 522;

        static string[] queries = { "Human resources", "Accountant",  "Secretary", "Entrepreneur", "Business development", "Consulting", "Marketing", "Business", "Finance", "Economics",
            "Machine Learning", "Data Science", "Web Developer", "Software Engineer", "Computer Science", "Computer Communication Engineering", "banking", "Sales Management", "Retails"  };
        
        static string[] universities = { "American University of Beirut", "Lebanese American University", "Université Saint-Joseph de Beyrouth", "Beirut Arab University", "Université Libanaise", "Notre Dame University",
            "Université Saint-Esprit de Kaslik", "University of Balamand", "Lebanese International University", "Haigazian University" };

        static string[] proxies = { "11.456.448.110:8080", "178.135.69.159", "89.108.147.203:1080", "89.108.165.40:1080", "194.126.8.90:80", "195.112.197.19:41021", "185.103.13.155:4153", "185.176.145.194:57746", "185.104.252.10:9090" };

        static void Main(string[] args)
        {
            //driver = new ChromeDriver(ChromDriverDirectory);
            //LoginLinkedin();
            //OpenExcel();
            //ScrapeData();
            //CloseExcel();

            //var excelCleaner = new ExcelCleaner("data.xlsx");
            //excelCleaner.Clean();
            //excelCleaner.Save();
            //excelCleaner.Dispose();
        }

        static void SearchGoogle()
        {
            var links = new List<string>();

            int sleepduration = 3000, sleepdurationthreshold = 3000;
            int botdetection = 1, botdetectionthreshold = 10;
            for (int q = 14; q < queries.Length; q++)
            {
                int start = 0;
                for (int u = start; u < universities.Length; u++)
                {
                    int oldlinkslength = links.Count;
                    var query = $"site:linkedin.com/in/ and \"{queries[q]}\" and \"{universities[u]}\"";
                    Console.WriteLine(query);
                    links = links.Concat(SearchGoogle(query)).ToList();

                    if (oldlinkslength == links.Count)
                    {
                        sleepduration = Math.Min(sleepduration * 5, 180000);
                        botdetection++;
                        u--;

                        if (botdetection == botdetectionthreshold)
                        {
                            return;
                        }
                    }
                    else
                    {
                        sleepduration = sleepdurationthreshold;
                        botdetection = 1;
                    }

                    Thread.Sleep(sleepduration);
                }
                Console.WriteLine();
            }
        }

        static void LoginLinkedin()
        {
            driver.Url = "https://www.linkedin.com";

            driver.FindElementById("session_key").SendKeys(email);
            driver.FindElementById("session_password").SendKeys(password);
            driver.FindElementByClassName("sign-in-form__submit-button").Click();
        }

        static List<string> SearchGoogle(string q)
        {
            driver.Url = $"https://www.google.com/search?q={q}";
            //var searchQuery = driver.FindElementByName("q");
            //searchQuery.SendKeys(q);
            //searchQuery.SendKeys(Keys.Return);

            var anchors = driver.FindElementsByTagName("a");
            var links = anchors
                .Where(anc => anc.GetAttribute("href") != null && anc.GetAttribute("href").StartsWith("https://lb.linkedin.com/in/"))
                .Select(anc => anc.GetAttribute("href"))
                .ToList();

            foreach (var link in links)
            {
                Console.WriteLine(link);
            }

            return links;
        }

        static void ScrapeData()
        {
            for (int i = offset; i < Constants.Links.Length; i++)
            {
                if (i - offset == 50)
                    break;

                try
                {
                    var link = Constants.Links[i];

                    int scrapingSleepDuration = 0;
                    var person = ScrapeData(link, scrapingSleepDuration);

                    if (person == null) continue;

                    if (person.isEmpty())
                    {
                        Console.WriteLine(i);
                        break;
                    }

                    if (People.FirstOrDefault(p => p.Name == person.Name) == null)
                    {
                        People.Add(person);
                        person.WriteToExcel(sheet, row: i + 2 - offset);

                        SaveExcel();
                    }
                }
                catch (Exception e) { }
            }
        }

        static Person ScrapeData(string link, int sleepDuration)
        {
            driver.Url = link;
            Thread.Sleep(2000);
            driver.ExecuteScript("scroll(0,5000)");
            var html = driver.PageSource;

            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            
            Thread.Sleep(sleepDuration / 10);

            string name = "";
            string job_title = "";
            string location = "";
            string company = "";
            string college = "";

            try
            {
                var left_panel_info = driver.FindElementByClassName("pv-text-details__left-panel").Text.Split("\r\n");
                name = left_panel_info[0];
                job_title = left_panel_info[1];
                location = left_panel_info[2];
            }
            catch (Exception) { }
            Thread.Sleep(sleepDuration / 10);

            try
            {
                var right_panel_info = driver.FindElementByClassName("pv-text-details__right-panel").Text.Split("\r\n");
                company = right_panel_info[0];
                college = right_panel_info[1];
            }
            catch (Exception) {  }

            var linkedin_link = link;

            Thread.Sleep(sleepDuration/10);

            // Experience
            try
            {
                driver.FindElementByXPath("//section[@class=\"pv-profile-section experience-section ember-view\"]/*/button").Click();
            }
            catch (Exception) { }

            var jobs = new List<string>();
            var experience_duration = "";
            try
            {
                var experienceSection = driver.FindElementById("experience-section");
                var jobsInfo = experienceSection.FindElements(By.XPath("//li/section/div/div/a/div[2]"));
                jobs = jobsInfo.Select(j => j.Text.Split("\r\n")[0] + " - " + j.Text.Split("\r\n")[2] + " - " + j.Text.Split("\r\n")[6]).ToList();

                experience_duration += jobsInfo.Aggregate("", (prev, curr) => prev + curr.Text.Split("\r\n")[6] + "\n");

                Thread.Sleep(sleepDuration/10);
                var moreJobsInfo = experienceSection.FindElements(By.XPath("//li/section/ul/li"));
                jobs = jobs.Concat(
                        moreJobsInfo
                            .Select(j => j.Text.Split("\r\n")[1] + " - " + experienceSection.Text.Split("\r\n")[2] + " - " + j.Text.Split("\r\n")[6])
                            .ToList())
                    .ToList();

                experience_duration += moreJobsInfo.Aggregate("", (prev, curr) => prev + curr.Text.Split("\r\n")[6] + "\n");
            }
            catch (Exception) { }
            Thread.Sleep(sleepDuration / 10);

            // Education
            var education = new List<string>();
            try
            {
                var educationSection = driver.FindElementById("education-section");
                var educationInfo = educationSection.FindElements(By.XPath("//ul/li/div/*/a/div[2]"));
                var date = educationInfo[0].Text.Split("\r\n")[6];
                var dates = date.Split("–");
                var int_date = int.Parse(dates[1]);
                if (int_date < graduationYearThreshold)
                    return null;

                education = educationInfo.Select(j => j.Text.Split("\r\n")[0] + " - " + j.Text.Split("\r\n")[4] + " - " + j.Text.Split("\r\n")[6]).ToList();
            }
            catch (Exception e) { }
            Thread.Sleep((int)0.1 * sleepDuration);

            // Certificates
            var certificates = new List<string>();
            try
            {
                var certificationsSection = driver.FindElementById("certifications-section");
                var certificatesInfo = certificationsSection.FindElements(By.XPath("//li/div/a/div[2]"));
                certificates = certificatesInfo.Select(j => j.Text.Split("\r\n")[0] + " - " + j.Text.Split("\r\n")[2]).ToList();
            }
            catch (Exception) { }
            Thread.Sleep((int)0.1 * sleepDuration);

            // Volunteering
            try
            {
                driver.FindElementByXPath("//section[contains(@class, 'volunteer-container')]/button").Click();
            }
            catch (Exception) { }
            Thread.Sleep(sleepDuration / 10);

            var volunteeringSections = driver.FindElementsByClassName("pv-volunteering-entity");
            var volunteeringActivities = new List<string>();
            var volunteering_duration = "";
            Thread.Sleep(sleepDuration / 10);
            try
            {
                foreach (var vsection in volunteeringSections)
                {
                    var activity = vsection.Text.Substring(0, vsection.Text.IndexOf("\r\nDates"));
                    activity = activity.Replace("Company Name", "-");
                    activity = activity.Replace("\r\n", " ");
                    volunteeringActivities.Add(activity);
                    volunteering_duration += vsection.Text.Split("\r\n")[6] + "\n";
                }
            }
            catch (Exception) { }

            //// Skills
            //try
            //{
            //    driver.FindElementByXPath("//section[contains(@class, \"pv-skill-categories-section\")]/*/button").Click();
            //}
            //catch (Exception) {  }

            //var skills = new List<string>();
            //var skillsSections = driver.FindElementsByClassName("pv-skill-category-entity__name");
            //skills = skillsSections.Select(s => s.Text.Split("\r\n")[0]).ToList();

            // Courses
            //var courses = driver.FindElementById("courses-expandable-content").Text;

            // Number of Projects
            string projects = "";
            try
            {
                projects = driver.FindElementByXPath("//section[contains(@class, \"pv-accomplishments-block projects\")]/h3/span[2]").Text;
            }
            catch (Exception) {  }
            Thread.Sleep(sleepDuration / 10);

            // Number of Languages
            string languages = "";
            try
            {
                languages = driver.FindElementByXPath("//section[contains(@class, 'pv-accomplishments-block languages')]/h3/span[2]").Text;
            }
            catch (Exception) { }
            Thread.Sleep(sleepDuration / 10);

            // Number of connections
            string connections = "";
            try
            {
                connections = driver.FindElementByXPath("//*[contains(@class, 'pv-top-card--list')]/li").Text.Replace("connections", "");
            }
            catch (Exception) { }
            //var driver.Find
            Thread.Sleep(sleepDuration / 10);

            string recommendations = "0";
            try
            {
                recommendations = driver.FindElementByXPath("//section[contains(@class, 'pv-recommendations-section')]//button[1]")
                    .Text.Replace("Received (", "").Replace(")", "");
            }
            catch (Exception) {  }

            return new Person
            {
                Name = name,
                JobTitle = job_title,
                Location = location,
                Company = company,
                College = college,
                LinkedinLink = linkedin_link,
                Projects = projects,
                Languages = languages,
                JobList = jobs,
                ExperienceDuration = experience_duration,
                EducationList = education,
                CertificationList = certificates,
                VolunteeringList = volunteeringActivities,
                VolunteeringDuration = volunteering_duration,
                Connections = connections,
                Recommendations = recommendations
                //SkillsList = skills
            };
        }
        static void OpenExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            excelPkg = new ExcelPackage(new FileInfo(OutputFileName));
            sheet = excelPkg.Workbook.Worksheets.Add(sheetName);
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

        static void SaveExcel()
        {
            excelPkg.Save();
        }

        static void CloseExcel()
        {
            sheet.Protection.IsProtected = false;
            sheet.Protection.AllowSelectLockedCells = false;
            excelPkg.Save();
            excelPkg.Dispose();
        }
    }
}
