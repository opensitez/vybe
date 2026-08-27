// vybe-test: csharp/linq_lambdas/linq_firstordefault_empty
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var nums = new List<int>();
__P((nums.FirstOrDefault()).ToString());
__Check("0");

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
