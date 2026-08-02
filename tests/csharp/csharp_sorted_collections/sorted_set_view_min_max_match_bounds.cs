// vybe-test: csharp/csharp_sorted_collections/sorted_set_view_min_max_match_bounds
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4, 5, 6 }; var view = ss.GetViewBetween(2, 5); __Check((view.Min).ToString(), "2"); __Check((view.Max).ToString(), "5");
