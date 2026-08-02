// vybe-test: csharp/csharp_list_dictionary/list_empty_initial_count_is_zero
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int>(); __Check((list.Count).ToString(), "0");
