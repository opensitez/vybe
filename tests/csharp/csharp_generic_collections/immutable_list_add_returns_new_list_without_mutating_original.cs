// vybe-test: csharp/csharp_generic_collections/immutable_list_add_returns_new_list_without_mutating_original
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

using static __Harness;

var original = System.Collections.Immutable.ImmutableList.Create(1,2,3);
var extended = original.Add(4);
__P((original.Count).ToString());
__P((extended.Count).ToString());
__Check("3\n4");

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
