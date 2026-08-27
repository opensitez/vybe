// vybe-test: csharp/csharp_lock_monitor/lock_decrement_inside_body_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object gate = new object();
int counter = 5;
lock (gate) { counter--; }
__P((counter).ToString());
__Check("4");

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
