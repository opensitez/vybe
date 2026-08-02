// vybe-test: csharp/csharp_list_dictionary/list_add_double_stores_fractional_number
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<double>(); list.Add(2.5); __Check((list[0]).ToString(), "2.5");
