// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_with_explicit_default_for_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); __Check((map.GetValueOrDefault("nope", -1)).ToString(), "-1");
