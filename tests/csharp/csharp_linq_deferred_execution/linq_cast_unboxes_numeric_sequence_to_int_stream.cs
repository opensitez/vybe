// vybe-test: csharp/csharp_linq_deferred_execution/linq_cast_unboxes_numeric_sequence_to_int_stream
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
object[] boxed = { 1, 2, 3 };
foreach (var value in boxed.Cast<int>()) Console.WriteLine(value + 1);
