// vybe-test: csharp/csharp_map_set_collections/dictionary_try_get_value_reports_missing_key
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); __Check((map.TryGetValue("a", out var value)).ToString(), "False");
