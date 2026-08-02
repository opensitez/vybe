// vybe-test: csharp/csharp_list_dictionary/dictionary_add_integer_key_stores_string
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string>(); map.Add(1, "one"); __Check((map[1]).ToString(), "one");
