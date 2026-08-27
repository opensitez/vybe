// vybe-test: csharp/csharp_exception_finally_control_flow/try_without_catch_still_executes_finally_when_body_throws
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

string trace = "";
try {
    try {
        throw new Exception("fail");
    } finally {
        trace += "finally;";
    }
}
catch (Exception) {
    trace += "handled;";
}
__P((trace).ToString());
__Check("finally;handled;");

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
