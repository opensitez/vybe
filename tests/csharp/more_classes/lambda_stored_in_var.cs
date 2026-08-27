// vybe-test: csharp/more_classes/lambda_stored_in_var
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

__P("Valid_lambda_stored_in_var");
__Check("Valid_lambda_stored_in_var");
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
