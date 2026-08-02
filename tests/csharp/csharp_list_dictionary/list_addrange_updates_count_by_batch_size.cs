// vybe-test: csharp/csharp_list_dictionary/list_addrange_updates_count_by_batch_size
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int>(); list.AddRange(new int[] { 4, 5, 6 }); __Check((list.Count).ToString(), "3");
