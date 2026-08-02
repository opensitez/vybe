// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_existing_key_returns_stored_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["hit"] = 33 }; __Check((map.GetValueOrDefault("hit")).ToString(), "33");
