// vybe-test: csharp/csharp_list_dictionary/list_nested_three_deep_reaches_innermost_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var outer = new List<List<List<int>>>(); var mid = new List<List<int>>(); var inner = new List<int> { 5 }; mid.Add(inner); outer.Add(mid); __Check((outer[0][0][0]).ToString(), "5");
