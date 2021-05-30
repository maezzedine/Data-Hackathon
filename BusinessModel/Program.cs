using Models;
using System;
using System.Collections.Generic;

namespace BusinessModel
{
    class Program
    {
        static Model model;
        static void Main(string[] args)
        {
            model = new Model();

            var fields = new List<Field> { Field.Data_Science, Field.Machine_Learning, Field.Computer_Science };
            Query(fields);
        }

        static void Query(List<Field> fields)
        {
            var recommendations = model.Query(fields);

            Console.WriteLine("Recommendations:");
            Console.WriteLine("----------------\n");

            foreach (var recommendation in recommendations)
            {
                Console.WriteLine(recommendation);
                Console.WriteLine();
            }
        }
    }
}
