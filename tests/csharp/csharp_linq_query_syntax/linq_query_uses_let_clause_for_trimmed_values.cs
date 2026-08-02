// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_let_clause_for_trimmed_values
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var raw = new[] { "  alpha  ", " beta", "gamma " };
var query = from value in raw
            let trimmed = value.Trim()
            select trimmed + ":" + trimmed.Length;
foreach (var item in query) Console.WriteLine(item);
