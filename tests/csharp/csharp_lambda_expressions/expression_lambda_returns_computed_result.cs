// vybe-test: csharp/csharp_lambda_expressions/expression_lambda_returns_computed_result
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

System.Func<int,int> f = x => x*x;
__P((f(5)).ToString());
__Check("25");

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
