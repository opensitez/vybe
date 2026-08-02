// vybe-test: csharp/csharp_extension_methods/extension_method_can_compose_with_generator_output
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using Demo; using System.Collections.Generic; namespace Demo { public static class NumberExt { public static IEnumerable<int> Twice(this IEnumerable<int> values) { foreach (var value in values) yield return value * 2; } } } foreach (var value in new[] { 1, 2 }.Twice()) Console.WriteLine(value);
