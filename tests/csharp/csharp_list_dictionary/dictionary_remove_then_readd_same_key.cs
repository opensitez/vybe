// vybe-test: csharp/csharp_list_dictionary/dictionary_remove_then_readd_same_key
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map.Remove("k"); map["k"] = 2; __Check((map["k"]).ToString(), "2");
