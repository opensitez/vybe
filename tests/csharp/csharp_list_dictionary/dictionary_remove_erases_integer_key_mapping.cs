// vybe-test: csharp/csharp_list_dictionary/dictionary_remove_erases_integer_key_mapping
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string> { [5] = "five" }; map.Remove(5); __Check((map.ContainsKey(5)).ToString(), "False");
