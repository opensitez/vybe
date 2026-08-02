// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_multiple_from_clauses_to_flatten_pairs
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var prefixes = new[] { "A", "B" };
var suffixes = new[] { 1, 2, 3 };
var query = from prefix in prefixes
            from suffix in suffixes
            where suffix != 2
            select prefix + suffix;
foreach (var item in query) Console.WriteLine(item);
