// vybe-test: csharp/linq_lambdas/linq_todictionary
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var names = new List<string> { "Alice", "Bob" }
;
var dict = names.ToDictionary(n => n, n => n.Length);
__P((dict["Alice"]).ToString());
__P((dict["Bob"]).ToString());
__Check("5\n3");

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
