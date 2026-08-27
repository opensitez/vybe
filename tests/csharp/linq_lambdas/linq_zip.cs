// vybe-test: csharp/linq_lambdas/linq_zip
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var names = new List<string> { "Alice", "Bob", "Charlie" }
;
var ages = new List<int> { 30, 25, 35 }
;
var pairs = names.Zip(ages, (n, a) => n + "=" + a).ToList();
foreach (var p in pairs) __P((p).ToString());
__Check("Alice=30\nBob=25\nCharlie=35");

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
