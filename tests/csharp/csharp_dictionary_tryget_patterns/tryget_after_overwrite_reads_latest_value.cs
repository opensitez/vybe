// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_overwrite_reads_latest_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 9; map.TryGetValue("k", out int v); __Check((v).ToString(), "9");
