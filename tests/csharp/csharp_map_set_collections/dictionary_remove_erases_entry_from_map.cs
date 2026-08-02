// vybe-test: csharp/csharp_map_set_collections/dictionary_remove_erases_entry_from_map
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 9 }; map.Remove("x"); __Check((map.ContainsKey("x")).ToString(), "False");
