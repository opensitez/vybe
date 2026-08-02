// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_on_dictionary_with_three_entries
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2, ["c"] = 3 }; map.TryGetValue("b", out int v); __Check((v).ToString(), "2");
