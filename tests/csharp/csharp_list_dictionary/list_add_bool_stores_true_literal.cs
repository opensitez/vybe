// vybe-test: csharp/csharp_list_dictionary/list_add_bool_stores_true_literal
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<bool>(); list.Add(true); __Check((list[0]).ToString(), "True");
