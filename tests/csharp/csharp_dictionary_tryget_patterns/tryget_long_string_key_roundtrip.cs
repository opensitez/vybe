// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_long_string_key_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["longer_name"] = 55 }; map.TryGetValue("longer_name", out int v); __Check((v).ToString(), "55");
