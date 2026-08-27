// vybe-test: csharp/exceptions_advanced/catch_when_filter
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new Exception("error 42");
}
catch (Exception e) when (e.Message.Contains("42")) {
    __P(("filtered catch: " + e.Message).ToString());
}
__Check("filtered catch: error 42");

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
