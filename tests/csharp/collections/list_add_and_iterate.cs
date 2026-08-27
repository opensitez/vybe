// vybe-test: csharp/collections/list_add_and_iterate
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var list = new List<string>();
list.Add("a");
list.Add("b");
list.Add("c");
foreach (var item in list) {
            __P((item).ToString());
        }
__Check("a\nb\nc");

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
