// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_distinguishes_similar_string_keys
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["cat"] = 3 }; __Check((map.TryGetValue("cat", out int v)).ToString(), "True"); __Check((map.TryGetValue("cats", out int w)).ToString(), "False");
