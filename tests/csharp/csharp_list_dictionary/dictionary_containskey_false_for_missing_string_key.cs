// vybe-test: csharp/csharp_list_dictionary/dictionary_containskey_false_for_missing_string_key
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 3 }; __Check((map.ContainsKey("z")).ToString(), "False");
