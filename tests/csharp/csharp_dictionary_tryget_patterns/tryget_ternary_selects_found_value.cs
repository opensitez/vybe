// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_ternary_selects_found_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 7 }; int result = map.TryGetValue("a", out int v) ? v : -1; __Check((result).ToString(), "7");
