// vybe-test: csharp/csharp_list_dictionary/list_add_zero_as_valid_element
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int>(); list.Add(0); __Check((list[0]).ToString(), "0"); __Check((list.Contains(0)).ToString(), "True");
