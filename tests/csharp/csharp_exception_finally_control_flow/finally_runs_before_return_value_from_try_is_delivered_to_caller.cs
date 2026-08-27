// vybe-test: csharp/csharp_exception_finally_control_flow/finally_runs_before_return_value_from_try_is_delivered_to_caller
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

int Pick() {
    try {
        return 2;
    } finally {
        __P(("cleanup").ToString());
    }
}
__P((Pick()).ToString());
__Check("cleanup\n2");

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
