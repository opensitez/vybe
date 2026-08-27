// vybe-test: csharp/csharp_exception_finally_control_flow/break_out_of_loop_runs_finally_before_exiting
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

string trace = "";
for (int i = 0; i < 3; i++) {
    try {
        trace += "body;";
        break;
    } finally {
        trace += "cleanup;";
    }
}
trace += "after";
__P((trace).ToString());
__Check("body;cleanup;after");

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
