// vybe-test: csharp/collections_advanced/list_removeat
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var list = new List<int> { 10, 20, 30, 40 }
;
list.RemoveAt(1);
foreach (var x in list) __P((x).ToString());
__Check("10\n30\n40");

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
