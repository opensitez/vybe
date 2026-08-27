// vybe-test: csharp/csharp_timer/threading_timer_fires_callback_after_delay
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

using static __Harness;

__P("Valid_threading_timer_fires_callback_after_delay");
__Check("Valid_threading_timer_fires_callback_after_delay");
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
