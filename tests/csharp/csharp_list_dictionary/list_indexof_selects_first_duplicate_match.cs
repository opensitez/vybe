// vybe-test: csharp/csharp_list_dictionary/list_indexof_selects_first_duplicate_match
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 7, 3, 7 }; __Check((list.IndexOf(7)).ToString(), "0");
