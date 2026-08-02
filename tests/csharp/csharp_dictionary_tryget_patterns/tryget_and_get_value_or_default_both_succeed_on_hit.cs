// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_and_get_value_or_default_both_succeed_on_hit
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["z"] = 44 }; bool ok = map.TryGetValue("z", out int t); int g = map.GetValueOrDefault("z"); __Check((ok).ToString(), "True"); __Check((g).ToString(), "44");
