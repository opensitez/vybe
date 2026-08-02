// vybe-test: csharp/csharp_sorted_collections/sorted_set_union_with_combines_sorted_uniques
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new SortedSet<int> { 1, 3 }; a.UnionWith(new[] { 2, 3, 4 }); __Check((a.Count).ToString(), "4"); __Check((a.Min).ToString(), "1"); __Check((a.Max).ToString(), "4");
