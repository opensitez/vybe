// vybe-test: csharp/csharp_list_dictionary/dictionary_remove_reduces_count_by_one
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; map.Remove("a"); __Check((map.Count).ToString(), "1");
