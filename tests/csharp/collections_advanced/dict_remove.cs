// vybe-test: csharp/collections_advanced/dict_remove
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 }, { "c", 3 } }
;
dict.Remove("b");
__P((dict.Count).ToString());
__P((dict.ContainsKey("b")).ToString());
__Check("2\nFalse");

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
