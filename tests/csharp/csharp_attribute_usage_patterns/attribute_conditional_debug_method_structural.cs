// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_debug_method_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public static void Trace(string m){Console.WriteLine(m);} public static void Run(){Trace("skip"); Console.WriteLine("seen");}} Log.Run();
