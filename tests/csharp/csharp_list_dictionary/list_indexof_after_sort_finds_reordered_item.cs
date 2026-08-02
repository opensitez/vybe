// vybe-test: csharp/csharp_list_dictionary/list_indexof_after_sort_finds_reordered_item
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 3, 1, 2 }; list.Sort(); __Check((list.IndexOf(2)).ToString(), "1");
