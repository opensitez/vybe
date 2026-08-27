// vybe-test: csharp/csharp_lambda_expressions/lambda_with_no_parameters_using_empty_parens
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

System.Func<string> greeting = () => "hello";
__P((greeting()).ToString());
__Check("hello");

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
