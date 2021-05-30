using Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessModel
{
    public class Model
    {
        public readonly List<Person> People;

        private const string filePath = "cleaned_data.xlsx";
        private const int numberOfRows = 1262;

        private List<string> DisregarderdCompanies = new List<string> { "company name" };

        public Model()
        {
            var reader = new ExcelReader(filePath);
            People = reader.GetPeople(numberOfRows);
        }

        public List<Recommendation> Query(List<Field> fields)
        {
            var recommendations = new List<Recommendation>();

            //var candidates = People.Where(p => p.College == university.ToString().Replace("_", " "));
            var companiesCounter = new Dictionary<string, int>();
            var ngosCounter = new Dictionary<string, int>();
            int LanguagesCounter = 0, CertificationsCounter = 0, ProjectsCounter = 0;
            int LanguagesTotal = 0, CertificationsTotal = 0, ProjectsTotal = 0;

            foreach (var field in fields)
            {
                var candidates = People.Where(p => 
                    p.JobTitle.Split(" - ")[0].ToLower().Trim() == field.ToString().Replace("_", " ").ToLower().Trim()
                );

                int activitiesMonthsCounter = 0;
                int candidatesWithActivitiesCounter = 0;

                foreach (var candidate in candidates)
                {
                    // Counting companies
                    var companies = candidate.GetCompanies();
                    foreach (var company in companies)
                    {
                        if (DisregarderdCompanies.Contains(company.ToLower().Trim()))
                        {
                            continue;
                        }

                        if (companiesCounter.ContainsKey(company))
                            companiesCounter[company]++;
                        else
                            companiesCounter[company] = 1;
                    }

                    // Counting NGOs
                    var ngos = candidate.GetNgos();
                    foreach (var ngo in ngos)
                    {
                        if (ngosCounter.ContainsKey(ngo))
                            ngosCounter[ngo]++;
                        else
                            ngosCounter[ngo] = 1;
                    }

                    // Counting Axtracurricular activities
                    if (int.TryParse(candidate.VolunteeringDuration, out int volunteeringMonths) && volunteeringMonths != 0)
                    {
                        activitiesMonthsCounter += volunteeringMonths;
                        candidatesWithActivitiesCounter++;
                    }

                    // Counting Languages
                    if (int.TryParse(candidate.Languages?.Trim(), out int languages))
                    {
                        LanguagesCounter += languages;
                        LanguagesTotal++;
                    }

                    // Counting Certifications
                    if (candidate.CertificationList != null && candidate.CertificationList.Count() - 2 != 0)
                    {
                        CertificationsCounter += candidate.CertificationList.Count() - 2;
                        CertificationsTotal++;
                    }

                    // Counting Projects
                    if (int.TryParse(candidate.Projects?.Trim(), out int projects))
                    {
                        ProjectsCounter += projects;
                        ProjectsTotal++;
                    }
                }

                if (candidatesWithActivitiesCounter != 0)
                {
                    var averageActivitiesMonths = activitiesMonthsCounter / candidatesWithActivitiesCounter;

                    recommendations.Add(new Recommendation
                    {
                        Field = field,
                        FieldValid = true,
                        Title = "Average amount of extracurricular activities (in months) per candidate",
                        Counter = averageActivitiesMonths,
                        CounterValid = true
                    });
                }
            }

            var companiesCounterPairList = companiesCounter.ToList();
            companiesCounterPairList.Sort((p1, p2) => p2.Value.CompareTo(p1.Value));

            var ngosCounterPairList = ngosCounter.ToList();
            ngosCounterPairList.Sort((p1, p2) => p2.Value.CompareTo(p1.Value));

            if (companiesCounterPairList.Count != 0)
            {
                recommendations.Add(new Recommendation
                {
                    Title = "Top mentioned companies to get experience in",
                    NameCounterList = companiesCounterPairList.Take(3).ToList(),
                });
            }

            if (ngosCounterPairList.Count != 0)
            {
                recommendations.Add(new Recommendation
                {
                    Title = "Top mentioned NGO to volunteer in",
                    NameCounterList = ngosCounterPairList.Take(3).ToList(),
                });
            }

            if (LanguagesTotal != 0)
            {
                var averageLanguages = Math.Min(1, LanguagesCounter / LanguagesTotal);
                recommendations.Add(new Recommendation
                {
                    Title = "Average number of languages spoken per candidate",
                    Counter = averageLanguages,
                    CounterValid = true
                });
            }

            if (ProjectsTotal != 0)
            {
                var averageProjects = ProjectsCounter / ProjectsTotal;
                recommendations.Add(new Recommendation
                {
                    Title = "Average number of projects completed per candidate",
                    Counter = averageProjects,
                    CounterValid = true
                });
            }

            if (CertificationsTotal != 0)
            {
                var averageCertifications = CertificationsCounter / CertificationsTotal;
                recommendations.Add(new Recommendation
                {
                    Title = "Average number of certificates earned per candidate",
                    Counter = averageCertifications,
                    CounterValid = true
                });
            }

            return recommendations;
        }
    }
}
