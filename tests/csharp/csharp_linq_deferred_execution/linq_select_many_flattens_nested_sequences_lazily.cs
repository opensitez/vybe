// vybe-test: csharp/csharp_linq_deferred_execution/linq_select_many_flattens_nested_sequences_lazily
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
var query = new[] { "ab", "c" }.SelectMany(word => word.Select(ch => ch));
foreach (var ch in query) Console.WriteLine(ch);
