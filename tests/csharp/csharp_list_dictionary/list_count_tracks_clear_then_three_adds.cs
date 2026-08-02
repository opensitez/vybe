// vybe-test: csharp/csharp_list_dictionary/list_count_tracks_clear_then_three_adds
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Clear(); list.Add(1); list.Add(2); list.Add(3); __Check((list.Count).ToString(), "3");
