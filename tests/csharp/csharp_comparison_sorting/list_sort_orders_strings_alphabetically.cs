// vybe-test: csharp/csharp_comparison_sorting/list_sort_orders_strings_alphabetically
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using System.Collections.Generic; var list = new List<string> { "c", "a", "b" }; list.Sort(); foreach (var value in list) Console.WriteLine(value);
