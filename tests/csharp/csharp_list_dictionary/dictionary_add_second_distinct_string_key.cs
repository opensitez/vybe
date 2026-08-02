// vybe-test: csharp/csharp_list_dictionary/dictionary_add_second_distinct_string_key
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("a", 1); map.Add("b", 2); __Check((map["b"]).ToString(), "2");
