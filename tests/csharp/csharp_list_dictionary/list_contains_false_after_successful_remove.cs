// vybe-test: csharp/csharp_list_dictionary/list_contains_false_after_successful_remove
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 5 }; list.Remove(5); __Check((list.Contains(5)).ToString(), "False");
