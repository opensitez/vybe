// vybe-test: csharp/csharp_list_dictionary/list_indexer_get_reads_first_added_item
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int>(); list.Add(5); list.Add(6); __Check((list[0]).ToString(), "5");
