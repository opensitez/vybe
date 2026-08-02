// vybe-test: csharp/csharp_list_dictionary/dictionary_string_keys_support_multi_character_names
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["alpha"] = 1, ["beta"] = 2 }; __Check((map["alpha"]).ToString(), "1");
