// vybe-test: csharp/csharp_linq_deferred_execution/linq_take_short_circuits_without_visiting_entire_source
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
int visited = 0;
var query = Enumerable.Range(1, 100).Select(x => { visited++; return x; }).Take(2);
foreach (var value in query) Console.WriteLine(value);
Console.WriteLine(visited);
