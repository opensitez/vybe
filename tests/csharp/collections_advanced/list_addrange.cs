// vybe-test: csharp/collections_advanced/list_addrange
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var list = new List<int> { 1, 2, 3 }
;
list.AddRange(new int[] { 4, 5 });
__P((list.Count).ToString());
foreach (var x in list) __P((x).ToString());
__Check("5\n1\n2\n3\n4\n5");

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
