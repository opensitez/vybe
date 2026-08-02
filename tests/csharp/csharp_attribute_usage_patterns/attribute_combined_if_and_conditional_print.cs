// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_combined_if_and_conditional_print
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

#define VYBETEST_PRE
using System; using System.Diagnostics; class App{[Conditional("DEBUG")] static void Log(){} static void Main(){#if VYBETEST_PRE __Check(("pre").ToString(), "pre"); #endif Log(); __Check(("post").ToString(), "post");}} App.Main();
