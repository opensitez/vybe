// vybe-test: csharp/csharp_list_dictionary/dictionary_containskey_still_true_after_overwrite
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 2; __Check((map.ContainsKey("k")).ToString(), "True");
