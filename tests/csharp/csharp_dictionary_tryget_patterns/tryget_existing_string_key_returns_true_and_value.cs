// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_existing_string_key_returns_true_and_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["alpha"] = 10 }; __Check((map.TryGetValue("alpha", out int v)).ToString(), "True"); __Check((v).ToString(), "10");
