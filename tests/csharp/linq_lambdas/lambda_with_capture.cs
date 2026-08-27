// vybe-test: csharp/linq_lambdas/lambda_with_capture
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

int multiplier = 3;
Func<int, int> mul = x => x * multiplier;
__P((mul(10)).ToString());
__P((mul(7)).ToString());
__Check("30\n21");

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
