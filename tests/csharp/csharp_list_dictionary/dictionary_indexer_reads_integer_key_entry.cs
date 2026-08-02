// vybe-test: csharp/csharp_list_dictionary/dictionary_indexer_reads_integer_key_entry
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string> { [2] = "two" }; __Check((map[2]).ToString(), "two");
