// vybe-test: csharp/csharp_list_dictionary/list_add_after_remove_restores_length
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Remove(1); list.Add(3); __Check((list.Count).ToString(), "2");
