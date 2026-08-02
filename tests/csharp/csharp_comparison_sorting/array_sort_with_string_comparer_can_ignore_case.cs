// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_string_comparer_can_ignore_case
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

var values = new[] { "b", "A", "c" }; System.Array.Sort(values, System.StringComparer.OrdinalIgnoreCase); foreach (var value in values) Console.WriteLine(value);
