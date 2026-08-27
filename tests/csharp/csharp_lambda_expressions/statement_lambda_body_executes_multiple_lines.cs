// vybe-test: csharp/csharp_lambda_expressions/statement_lambda_body_executes_multiple_lines
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

System.Func<int,int> fact = null;
fact = n => { if(n<=1) return 1; return n*fact(n-1); }
;
__P((fact(5)).ToString());
__Check("120");

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
