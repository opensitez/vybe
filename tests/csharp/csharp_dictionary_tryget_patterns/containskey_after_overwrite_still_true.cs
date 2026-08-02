// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_after_overwrite_still_true
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 100; __Check((map.ContainsKey("k")).ToString(), "True");
