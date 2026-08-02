// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_second_distinct_add
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("first", 1); map.Add("second", 2); map.TryGetValue("second", out int v); __Check((v).ToString(), "2");
