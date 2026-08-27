// vybe-test: csharp/csharp_concurrency_sync/monitor_try_enter_returns_false_when_already_locked
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

using static __Harness;

__P("Valid_monitor_try_enter_returns_false_when_already_locked");
__Check("Valid_monitor_try_enter_returns_false_when_already_locked");
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
