// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_add_method_finds_added_pair
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("one", 1); __Check((map.TryGetValue("one", out int v)).ToString(), "True"); __Check((v).ToString(), "1");
