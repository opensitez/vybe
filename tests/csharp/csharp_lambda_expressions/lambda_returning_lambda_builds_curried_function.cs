// vybe-test: csharp/csharp_lambda_expressions/lambda_returning_lambda_builds_curried_function
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

System.Func<int,System.Func<int,int>> add = a => b => a+b;
var add5 = add(5);
__P((add5(3)).ToString());
__Check("8");

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
