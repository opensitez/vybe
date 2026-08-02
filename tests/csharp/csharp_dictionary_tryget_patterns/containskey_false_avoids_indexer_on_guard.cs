// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_false_avoids_indexer_on_guard
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); if (!map.ContainsKey("z")) __Check(("skip").ToString(), "skip");
