// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_and_tryget_agree_on_present_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 8 }; __Check((map.ContainsKey("k")).ToString(), "True"); __Check((map.TryGetValue("k", out int v)).ToString(), "True");
