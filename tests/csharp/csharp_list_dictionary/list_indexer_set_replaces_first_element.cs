// vybe-test: csharp/csharp_list_dictionary/list_indexer_set_replaces_first_element
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<string> { "old", "keep" }; list[0] = "new"; __Check((list[0]).ToString(), "new");
