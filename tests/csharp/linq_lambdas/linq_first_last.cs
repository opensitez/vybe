// vybe-test: csharp/linq_lambdas/linq_first_last
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var nums = new List<int> { 10, 20, 30 }
;
__P((nums.First()).ToString());
__P((nums.Last()).ToString());
__Check("10\n30");

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
