// vybe-test: csharp/csharp_sorted_collections/sorted_set_get_view_between_single_element
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 10, 20, 30 }; var view = ss.GetViewBetween(20, 20); __Check((view.Min).ToString(), "20");
