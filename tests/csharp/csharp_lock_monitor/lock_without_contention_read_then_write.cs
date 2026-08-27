// vybe-test: csharp/csharp_lock_monitor/lock_without_contention_read_then_write
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object gate = new object();
int counter = 1;
lock (gate) {
    int snapshot = counter;
    counter = snapshot + 4;
}
__P((counter).ToString());
__Check("5");

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
