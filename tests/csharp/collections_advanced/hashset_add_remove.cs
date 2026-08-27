// vybe-test: csharp/collections_advanced/hashset_add_remove
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var set = new HashSet<string>();
set.Add("apple");
set.Add("banana");
set.Add("apple");
__P((set.Count).ToString());
set.Remove("apple");
__P((set.Count).ToString());
__P((set.Contains("apple")).ToString());
__Check("2\n1\nFalse");

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
