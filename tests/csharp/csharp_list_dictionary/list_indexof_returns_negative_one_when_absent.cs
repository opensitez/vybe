// vybe-test: csharp/csharp_list_dictionary/list_indexof_returns_negative_one_when_absent
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2 }; __Check((list.IndexOf(99)).ToString(), "-1");
