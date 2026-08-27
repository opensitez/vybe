// vybe-test: csharp/csharp_immutable_collections/immutable_list_add_returns_new_list_old_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

using static __Harness;

var a=System.Collections.Immutable.ImmutableList<int>.Empty;
var b=a.Add(1).Add(2).Add(3);
__P((a.Count).ToString());
__P((b.Count).ToString());
__Check("0\n3");

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
