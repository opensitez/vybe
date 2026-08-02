// vybe-test: csharp/csharp_linq_materialization/linq_reverse_materializes_reversed_order_in_new_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

using System.Linq;
foreach (var value in new[] { 1, 2, 3 }.Reverse()) Console.WriteLine(value);
