// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_after_clear_is_false
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); __Check((map.ContainsKey("a")).ToString(), "False");
