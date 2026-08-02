// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_parallel_keys_reorders_items_by_key
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

var keys = new[] { 2, 1 }; var items = new[] { "b", "a" }; System.Array.Sort(keys, items); foreach (var value in items) Console.WriteLine(value);
