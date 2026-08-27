// vybe-test: csharp/csharp_exception_finally_control_flow/exception_rethrown_from_catch_is_handled_by_enclosing_try
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

string trace = "";
try {
    try {
        throw new Exception("first");
    } catch (Exception) {
        trace += "inner;";
        throw new Exception("second");
    }
}
catch (Exception e) {
    trace += "outer:" + e.Message;
}
__P((trace).ToString());
__Check("inner;outer:second");

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
