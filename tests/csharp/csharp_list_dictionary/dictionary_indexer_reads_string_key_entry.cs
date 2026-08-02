// vybe-test: csharp/csharp_list_dictionary/dictionary_indexer_reads_string_key_entry
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 4 }; __Check((map["k"]).ToString(), "4");
