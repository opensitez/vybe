// vybe-test: csharp/linq_lambdas/linq_min_max
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var nums = new List<int> { 5, 3, 8, 1, 9 }
;
__P((nums.Min()).ToString());
__P((nums.Max()).ToString());
__Check("1\n9");

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
