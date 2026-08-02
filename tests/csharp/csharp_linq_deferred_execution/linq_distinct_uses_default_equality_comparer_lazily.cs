// vybe-test: csharp/csharp_linq_deferred_execution/linq_distinct_uses_default_equality_comparer_lazily
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
foreach (var value in new[] { 1, 1, 2, 2, 3 }.Distinct()) Console.WriteLine(value);
