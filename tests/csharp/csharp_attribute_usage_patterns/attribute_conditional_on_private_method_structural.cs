// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_on_private_method_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; using System.Diagnostics; class S{[Conditional("DEBUG")] void Trace(){} public void Run(){Trace(); __Check(("ok").ToString(), "ok");}} new S().Run();
