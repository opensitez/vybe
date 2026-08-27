// vybe-test: csharp/collections_advanced/list_trueforall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var list = new List<int> { 2, 4, 6, 8 }
;
__P((list.TrueForAll(x => x % 2 == 0)).ToString());
list.Add(3);
__P((list.TrueForAll(x => x % 2 == 0)).ToString());
__Check("True\nFalse");

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
