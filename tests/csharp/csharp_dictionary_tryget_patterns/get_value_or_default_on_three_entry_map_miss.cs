// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_on_three_entry_map_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2, ["c"] = 3 }; __Check((map.GetValueOrDefault("d", 0)).ToString(), "0");
