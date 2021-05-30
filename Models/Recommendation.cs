using System.Collections.Generic;
using System.Text;

namespace Models
{
    public class Recommendation
    {
        public Field Field { get; set; }
        public bool FieldValid { get; set; }
        public List<KeyValuePair<string, int>> NameCounterList { get; set; }
        public string Title { get; set; }
        public int Counter { get; set; }
        public bool CounterValid { get; set; }

        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            if (FieldValid)
            {
                stringBuilder.AppendLine($"In the field of {Field.ToString().Replace("_", " ")}:");
            }

            stringBuilder.AppendLine(Title + ":");
            
            if (CounterValid)
            {
                stringBuilder.AppendLine(Counter.ToString());
            }

            if (NameCounterList != null)
            {
                stringBuilder.AppendLine();
                foreach (var pair in NameCounterList)
                {
                    stringBuilder.AppendLine($"{pair.Key}\t{pair.Value} candidates");
                }
            }

            return stringBuilder.ToString();
        }
    }
}
