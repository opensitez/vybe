// vybe-test: csharp/csharp_list_dictionary/dictionary_clear_resets_entry_count
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); __Check((map.Count).ToString(), "0");
