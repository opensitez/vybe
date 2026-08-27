// vybe-test: csharp/collections_advanced/list_findall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var list = new List<int> { 1, 2, 3, 4, 5, 6 }
;
var evens = list.FindAll(x => x % 2 == 0);
foreach (var x in evens) __P((x).ToString());
__Check("2\n4\n6");

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
