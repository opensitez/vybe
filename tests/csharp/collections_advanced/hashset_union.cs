// vybe-test: csharp/collections_advanced/hashset_union
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var a = new HashSet<int> { 1, 2, 3 }
;
var b = new HashSet<int> { 3, 4, 5 }
;
a.UnionWith(b);
__P((a.Count).ToString());
__Check("5");

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
