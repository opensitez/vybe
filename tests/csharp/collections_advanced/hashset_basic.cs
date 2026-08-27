// vybe-test: csharp/collections_advanced/hashset_basic
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var set = new HashSet<int> { 1, 2, 3, 2, 1 }
;
__P((set.Count).ToString());
__P((set.Contains(2)).ToString());
__P((set.Contains(5)).ToString());
__Check("3\nTrue\nFalse");

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
