// vybe-test: csharp/csharp_comparison_sorting/then_by_breaks_ties_after_primary_sort_key
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using System.Linq; var values = new[] { "ba", "aa", "c" }.OrderBy(text => text.Length).ThenBy(text => text); foreach (var value in values) Console.WriteLine(value);
