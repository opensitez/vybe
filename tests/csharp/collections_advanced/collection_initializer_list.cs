// vybe-test: csharp/collections_advanced/collection_initializer_list
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var names = new List<string> { "Alice", "Bob", "Charlie" }
;
__P((names.Count).ToString());
__P((names[1]).ToString());
__Check("3\nBob");

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
