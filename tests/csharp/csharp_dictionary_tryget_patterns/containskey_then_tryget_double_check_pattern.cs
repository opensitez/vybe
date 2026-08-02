// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_then_tryget_double_check_pattern
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["c"] = 9 }; int outVal = 0; if (map.ContainsKey("c") && map.TryGetValue("c", out outVal)) __Check((outVal).ToString(), "9");
