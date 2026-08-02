// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_distinguishes_similar_string_keys
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["ab"] = 1 }; __Check((map.ContainsKey("ab")).ToString(), "True"); __Check((map.ContainsKey("a")).ToString(), "False");
