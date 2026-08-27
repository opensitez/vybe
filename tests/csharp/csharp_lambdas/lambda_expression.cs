// vybe-test: csharp/csharp_lambdas/lambda_expression
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

using static __Harness;

var double_it = (int x) => x * 2;
__P((double_it(5)).ToString());
__Check("10");

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
