// vybe-test: csharp/csharp_dictionary_tryget_patterns/indexer_read_matches_tryget_for_same_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 4 }; map.TryGetValue("x", out int viaTry); __Check((map["x"] == viaTry).ToString(), "True");
