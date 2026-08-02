// vybe-test: csharp/csharp_comparison_sorting/list_sort_with_custom_comparison_orders_descending
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using System.Collections.Generic; var list = new List<int> { 1, 3, 2 }; list.Sort((left, right) => right.CompareTo(left)); foreach (var value in list) Console.WriteLine(value);
