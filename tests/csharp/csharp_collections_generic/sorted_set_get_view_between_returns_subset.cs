// vybe-test: csharp/csharp_collections_generic/sorted_set_get_view_between_returns_subset
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s=new System.Collections.Generic.SortedSet<int>{1,2,3,4,5};
var view=s.GetViewBetween(2,4);
__Check((view.Count).ToString(), "3");
