// vybe-test: csharp/collections_advanced/list_insert_at
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var list = new List<string> { "a", "c", "d" }
;
list.Insert(1, "b");
foreach (var s in list) __P((s).ToString());
__Check("a\nb\nc\nd");

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
