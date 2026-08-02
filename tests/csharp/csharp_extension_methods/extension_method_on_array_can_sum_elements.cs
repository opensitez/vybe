// vybe-test: csharp/csharp_extension_methods/extension_method_on_array_can_sum_elements
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using Demo; namespace Demo { public static class ArrayExt { public static int SumAll(this int[] values) { int total = 0; foreach (var value in values) total += value; return total; } } } Console.WriteLine(new[] { 1, 2, 3 }.SumAll());
