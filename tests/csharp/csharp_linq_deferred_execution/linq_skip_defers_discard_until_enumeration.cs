// vybe-test: csharp/csharp_linq_deferred_execution/linq_skip_defers_discard_until_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
var query = new[] { 10, 20, 30, 40 }.Skip(2);
foreach (var value in query) Console.WriteLine(value);
