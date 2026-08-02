// vybe-test: csharp/csharp_list_dictionary/list_indexof_zero_after_head_insert
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 2 }; list.Insert(0, 1); __Check((list.IndexOf(1)).ToString(), "0");
