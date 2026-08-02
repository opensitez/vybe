// vybe-test: csharp/csharp_comparison_sorting/order_by_with_key_projection_sorts_by_length
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using System.Linq; var values = new[] { "bbb", "a", "cc" }.OrderBy(text => text.Length); foreach (var value in values) Console.WriteLine(value);
