// vybe-test: csharp/csharp_list_dictionary/dictionary_int_key_overwrite_via_indexer
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [1] = 10 }; map[1] = 99; __Check((map[1]).ToString(), "99");
