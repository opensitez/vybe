// vybe-test: csharp/csharp_list_dictionary/list_nested_indexer_reaches_inner_list
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var outer = new List<List<int>> { new List<int> { 10, 20 } }; __Check((outer[0][1]).ToString(), "20");
