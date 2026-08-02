// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_twice_on_same_key_yields_same_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["dup"] = 6 }; map.TryGetValue("dup", out int a); map.TryGetValue("dup", out int b); __Check((a == b).ToString(), "True");
