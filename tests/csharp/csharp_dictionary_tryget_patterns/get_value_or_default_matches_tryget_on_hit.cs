// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_matches_tryget_on_hit
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["p"] = 12 }; map.TryGetValue("p", out int t); __Check((map.GetValueOrDefault("p") == t).ToString(), "True");
