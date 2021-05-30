using System.Text.RegularExpressions;

namespace ScrapingData
{
    public class TimeDuration
    {
        public int Months { get; set; }
        public int Years { get; set; }

        public TimeDuration(string input)
        {
            if (input == null)
            {
                return;
            }

            var times = input.Split("\n");
            foreach (var time in times)
            {
                var subtimes = time.Split();

                for (int i = 0; i < subtimes.Length - 1; i++)
                {
                    if (isNumber(subtimes[i]))
                    {
                        var unit = subtimes[i + 1].Trim();
                        if (unit == "yr" || unit == "yrs")
                        {
                            Years += int.Parse(subtimes[i]);
                        }
                        else if (unit == "mos" || unit == "mo") 
                        {
                            Months += int.Parse(subtimes[i]);
                        }
                    }
                }
            }
        }

        public int getMonths() => Years * 12 + Months;

        private bool isNumber(string s)
        {
            return Regex.IsMatch(s, "[0-9]");
        }
    }
}
