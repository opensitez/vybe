// vybe-test: csharp/csharp_error_handling/nested_try_catch
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

using static __Harness;

try {
    try {
        throw new Exception("inner");
    } catch (Exception e) {
        __P(("inner: " + e.Message).ToString());
        throw new Exception("rethrown");
    }
}
catch (Exception e) {
    __P(("outer: " + e.Message).ToString());
}
__Check("inner: inner\nouter: rethrown");

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
