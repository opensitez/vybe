// vybe-test: csharp/csharp_linq_query_syntax/linq_query_filters_even_numbers_then_projects_squares
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var values = new[] { 1, 2, 3, 4, 5, 6 };
var query = from value in values
            where value % 2 == 0
            select value * value;
foreach (var value in query) Console.WriteLine(value);
