// vybe-test: csharp/exceptions_advanced/nested_try_inner_uncaught_propagates
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    try {
        throw new InvalidOperationException("oops");
    } catch (ArgumentException) {
        __P(("wrong handler").ToString());
    }
}
catch (InvalidOperationException e) {
    __P(("outer got: " + e.Message).ToString());
}
__Check("outer got: oops");

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
