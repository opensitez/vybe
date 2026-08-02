// vybe-test: csharp/csharp_linq_deferred_execution/linq_where_filter_runs_only_during_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
int checks = 0;
var query = new[] { 1, 2, 3, 4 }.Where(x => { checks++; return x % 2 == 0; });
Console.WriteLine(checks);
foreach (var value in query) Console.WriteLine(value);
Console.WriteLine(checks);
