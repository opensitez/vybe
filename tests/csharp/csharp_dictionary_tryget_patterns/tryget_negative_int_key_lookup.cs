// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_negative_int_key_lookup
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [-1] = 100 }; map.TryGetValue(-1, out int v); __Check((v).ToString(), "100");
