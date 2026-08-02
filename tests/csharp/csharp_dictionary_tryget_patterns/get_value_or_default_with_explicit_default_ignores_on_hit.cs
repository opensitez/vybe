// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_with_explicit_default_ignores_on_hit
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["ok"] = 5 }; __Check((map.GetValueOrDefault("ok", 99)).ToString(), "5");
