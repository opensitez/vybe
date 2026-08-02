// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_int_key_miss_returns_false
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "a" }; __Check((map.TryGetValue(99, out string s)).ToString(), "False");
