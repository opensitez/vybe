// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_empty_string_key_is_valid
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { [""] = 1 }; map.TryGetValue("", out int v); __Check((v).ToString(), "1");
