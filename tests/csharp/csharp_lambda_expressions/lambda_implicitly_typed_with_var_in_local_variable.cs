// vybe-test: csharp/csharp_lambda_expressions/lambda_implicitly_typed_with_var_in_local_variable
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

var f = (int x) => x + 1;
__P((f(9)).ToString());
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
