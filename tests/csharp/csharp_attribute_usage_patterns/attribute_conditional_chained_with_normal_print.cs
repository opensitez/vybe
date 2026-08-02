// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_chained_with_normal_print
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using System; using System.Diagnostics; class P{[Conditional("DEBUG")] static void D(){Console.WriteLine("d");} static void N(){Console.WriteLine("n");} public static void Go(){D(); N();}} P.Go();
