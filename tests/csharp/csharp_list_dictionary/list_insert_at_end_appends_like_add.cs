// vybe-test: csharp/csharp_list_dictionary/list_insert_at_end_appends_like_add
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Insert(2, 3); __Check((list[2]).ToString(), "3");
