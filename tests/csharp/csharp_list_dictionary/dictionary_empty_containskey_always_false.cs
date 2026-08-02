// vybe-test: csharp/csharp_list_dictionary/dictionary_empty_containskey_always_false
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); __Check((map.ContainsKey("any")).ToString(), "False");
