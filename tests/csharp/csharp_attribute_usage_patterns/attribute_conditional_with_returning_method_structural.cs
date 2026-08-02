// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_with_returning_method_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using System; using System.Diagnostics; class Calc{[Conditional("DEBUG")] static void Log(int x){Console.WriteLine(x);} public static int Add(int a,int b){Log(a+b); return a+b;}} Console.WriteLine(Calc.Add(2,3));
