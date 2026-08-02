// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_does_not_affect_other_method
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public static void A(){} public static void B(){__Check(("b").ToString(), "b");} public static void Run(){A(); B();}} Log.Run();
