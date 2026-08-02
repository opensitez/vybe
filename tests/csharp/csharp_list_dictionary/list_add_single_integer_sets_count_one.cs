// vybe-test: csharp/csharp_list_dictionary/list_add_single_integer_sets_count_one
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int>(); list.Add(42); __Check((list.Count).ToString(), "1");
