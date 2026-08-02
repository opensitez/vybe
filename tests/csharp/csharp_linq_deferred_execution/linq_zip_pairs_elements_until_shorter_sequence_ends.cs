// vybe-test: csharp/csharp_linq_deferred_execution/linq_zip_pairs_elements_until_shorter_sequence_ends
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using System.Linq;
foreach (var pair in new[] { 1, 2, 3 }.Zip(new[] { 10, 20 }, (a, b) => a + b)) Console.WriteLine(pair);
