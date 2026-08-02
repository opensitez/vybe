// vybe-test: csharp/csharp_list_dictionary/dictionary_indexer_sets_brand_new_int_key
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string>(); map[42] = "answer"; __Check((map[42]).ToString(), "answer");
