// vybe-test: csharp/csharp_linq_deferred_execution/linq_oftype_filters_runtime_types_during_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
object[] items = { 1, "a", 2, "b", 3 };
foreach (var text in items.OfType<string>()) Console.WriteLine(text);
