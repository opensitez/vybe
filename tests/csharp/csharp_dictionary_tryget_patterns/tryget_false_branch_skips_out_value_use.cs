// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_false_branch_skips_out_value_use
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); if (!map.TryGetValue("x", out var val)) __Check(("miss").ToString(), "miss");
