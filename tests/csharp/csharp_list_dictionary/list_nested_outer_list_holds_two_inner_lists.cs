// vybe-test: csharp/csharp_list_dictionary/list_nested_outer_list_holds_two_inner_lists
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var outer = new List<List<int>>(); outer.Add(new List<int> { 1 }); outer.Add(new List<int> { 2, 3 }); __Check((outer.Count).ToString(), "2");
