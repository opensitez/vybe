// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_readd_after_remove_finds_new_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["r"] = 1 }; map.Remove("r"); map["r"] = 2; map.TryGetValue("r", out int v); __Check((v).ToString(), "2");
