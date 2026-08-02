// vybe-test: csharp/csharp_bcl_collections/hashset_union_with_merges_distinct_elements_from_other_set
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = new System.Collections.Generic.HashSet<int> { 1, 2 };
var right = new System.Collections.Generic.HashSet<int> { 2, 3 };
left.UnionWith(right);
__Check((left.Count).ToString(), "3");
