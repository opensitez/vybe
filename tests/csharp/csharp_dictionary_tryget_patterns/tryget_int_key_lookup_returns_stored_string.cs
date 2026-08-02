// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_int_key_lookup_returns_stored_string
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string> { [42] = "answer" }; map.TryGetValue(42, out string s); __Check((s).ToString(), "answer");
