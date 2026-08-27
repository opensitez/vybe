// vybe-test: csharp/linq_lambdas/linq_orderby_thenby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var names = new List<string> { "Charlie", "Alice", "Bob", "Alice" }
;
var sorted = names.OrderBy(n => n).ToList();
foreach (var n in sorted) __P((n).ToString());
__Check("Alice\nAlice\nBob\nCharlie");

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
