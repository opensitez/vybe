// vybe-test: csharp/csharp_linq_deferred_execution/linq_orderby_defers_sort_until_materialization
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
int comparisons = 0;
var query = new[] { 3, 1, 2 }.OrderBy(x => { comparisons++; return x; });
Console.WriteLine(comparisons);
foreach (var value in query) Console.WriteLine(value);
