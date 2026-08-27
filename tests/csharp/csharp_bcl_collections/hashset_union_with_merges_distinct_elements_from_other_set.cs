// vybe-test: csharp/csharp_bcl_collections/hashset_union_with_merges_distinct_elements_from_other_set
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

using static __Harness;

var left = new System.Collections.Generic.HashSet<int> { 1, 2 }
;
var right = new System.Collections.Generic.HashSet<int> { 2, 3 }
;
left.UnionWith(right);
__P((left.Count).ToString());
__Check("3");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
