// vybe-test: csharp/csharp_list_dictionary/list_reverse_twice_restores_original_order
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Reverse(); list.Reverse(); __Check((list[0]).ToString(), "1"); __Check((list[2]).ToString(), "3");
