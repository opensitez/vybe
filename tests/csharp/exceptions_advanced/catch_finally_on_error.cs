// vybe-test: csharp/exceptions_advanced/catch_finally_on_error
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

string result = "start";
try {
    int x = 10 / Math.Max(0, 0);
    result = "never";
}
catch (DivideByZeroException) {
    result = "caught";
}
finally {
    result += " + finally";
}
__P((result).ToString());
__Check("caught + finally");

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
