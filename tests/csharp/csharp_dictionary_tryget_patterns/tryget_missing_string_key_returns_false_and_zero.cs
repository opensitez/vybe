// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_missing_string_key_returns_false_and_zero
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); __Check((map.TryGetValue("ghost", out int v)).ToString(), "False"); __Check((v).ToString(), "0");
