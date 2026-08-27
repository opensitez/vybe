// vybe-test: csharp/csharp_collections_generic/sorted_set_get_view_between_returns_subset
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

using static __Harness;

var s=new System.Collections.Generic.SortedSet<int>{1,2,3,4,5}
;
var view=s.GetViewBetween(2,4);
__P((view.Count).ToString());
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
