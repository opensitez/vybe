// vybe-test: csharp/csharp_collections_generic/hash_set_union_with_intersects_and_except
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

using static __Harness;

var a=new System.Collections.Generic.HashSet<int>{1,2,3,4}
;
var b=new System.Collections.Generic.HashSet<int>{3,4,5,6}
;
a.IntersectWith(b);
__P((a.Count).ToString());
__P((a.Contains(3)).ToString());
__Check("2\nTrue");

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
