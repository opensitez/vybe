// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_clear_always_fails
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); __Check((map.TryGetValue("a", out int v)).ToString(), "False");
