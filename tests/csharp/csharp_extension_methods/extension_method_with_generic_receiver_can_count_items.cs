// vybe-test: csharp/csharp_extension_methods/extension_method_with_generic_receiver_can_count_items
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using Demo; using System.Collections.Generic; namespace Demo { public static class EnumerableExt { public static int CountItems<T>(this IEnumerable<T> items) { int total = 0; foreach (var _ in items) total++; return total; } } } Console.WriteLine(new[] { 1, 2, 3 }.CountItems());
