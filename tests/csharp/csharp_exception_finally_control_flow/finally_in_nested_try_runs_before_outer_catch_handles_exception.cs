// vybe-test: csharp/csharp_exception_finally_control_flow/finally_in_nested_try_runs_before_outer_catch_handles_exception
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

try {
    try {
        throw new Exception("boom");
    } finally {
        __P(("inner-finally").ToString());
    }
}
catch (Exception) {
    __P(("outer-catch").ToString());
}
__Check("inner-finally\nouter-catch");

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
