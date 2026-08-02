// vybe-test: csharp/csharp_list_dictionary/list_removeat_zero_drops_first_item
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; list.RemoveAt(0); __Check((list[0]).ToString(), "20");
