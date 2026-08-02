// vybe-test: csharp/csharp_linq_deferred_execution/linq_select_does_not_run_until_foreach_starts
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
int sideEffects = 0;
var query = new[] { 1, 2 }.Select(x => { sideEffects++; return x; });
Console.WriteLine(sideEffects);
foreach (var _ in query) { }
Console.WriteLine(sideEffects);
