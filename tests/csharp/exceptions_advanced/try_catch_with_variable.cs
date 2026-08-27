// vybe-test: csharp/exceptions_advanced/try_catch_with_variable
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

__P("Valid_try_catch_with_variable");
__Check("Valid_try_catch_with_variable");
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
