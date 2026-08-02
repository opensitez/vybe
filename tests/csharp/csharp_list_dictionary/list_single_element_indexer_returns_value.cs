// vybe-test: csharp/csharp_list_dictionary/list_single_element_indexer_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<string> { "solo" }; __Check((list[0]).ToString(), "solo");
