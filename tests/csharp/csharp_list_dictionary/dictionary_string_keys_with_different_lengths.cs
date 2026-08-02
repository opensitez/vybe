// vybe-test: csharp/csharp_list_dictionary/dictionary_string_keys_with_different_lengths
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["longer"] = 2 }; __Check((map["longer"]).ToString(), "2");
