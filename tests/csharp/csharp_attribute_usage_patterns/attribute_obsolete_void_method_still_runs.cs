// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_void_method_still_runs
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("legacy")] public void Ping(){__Check(("ping").ToString(), "ping");}} new S().Ping();
