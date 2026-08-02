// vybe-test: csharp/csharp_list_dictionary/dictionary_count_reflects_two_add_operations
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("a", 1); map.Add("b", 2); __Check((map.Count).ToString(), "2");
