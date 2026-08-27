// vybe-test: csharp/csharp_lock_monitor/lock_for_loop_accumulates_squares_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object gate = new object();
int counter = 0;
for (int i = 1; i <= 4; i++) {
    lock (gate) { counter += i * i; }
}
__P((counter).ToString());
__Check("30");

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
