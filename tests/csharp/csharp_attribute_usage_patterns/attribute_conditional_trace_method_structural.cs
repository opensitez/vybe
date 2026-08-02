// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_trace_method_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using System; using System.Diagnostics; class Log{[Conditional("TRACE")] public static void Mark(){Console.WriteLine("mark");} public static void Run(){Mark(); Console.WriteLine("after");}} Log.Run();
