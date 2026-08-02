// vybe-test: csharp/csharp_list_dictionary/list_clear_allows_subsequent_add
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1 }; list.Clear(); list.Add(7); __Check((list[0]).ToString(), "7");
