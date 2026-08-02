// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_on_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public void Trace(){Console.WriteLine("t");} public void Run(){Trace(); Console.WriteLine("r");}} new Log().Run();
