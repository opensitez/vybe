// vybe-test: csharp/csharp_sorted_collections/sorted_set_get_view_between_inclusive_range
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4, 5 }; var view = ss.GetViewBetween(2, 4); __Check((view.Count).ToString(), "3");
