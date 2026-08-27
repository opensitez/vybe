// vybe-test: csharp/exceptions_advanced/try_catch_basic
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    int x = 10 / Math.Max(0, 0);
}
catch (DivideByZeroException) {
    __P(("caught divide by zero").ToString());
}
__Check("caught divide by zero");

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
