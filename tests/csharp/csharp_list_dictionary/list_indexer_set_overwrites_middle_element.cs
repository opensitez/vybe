// vybe-test: csharp/csharp_list_dictionary/list_indexer_set_overwrites_middle_element
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list[1] = 99; __Check((list[1]).ToString(), "99");
