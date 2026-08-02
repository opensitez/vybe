// vybe-test: csharp/csharp_list_dictionary/list_addrange_from_int_array_extends_tail
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.AddRange(new int[] { 3, 4 }); __Check((list[3]).ToString(), "4");
