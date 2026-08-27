// vybe-test: csharp/csharp_lock_monitor/lock_separate_gates_no_interference
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object g1 = new object();
object g2 = new object();
int c1 = 0;
int c2 = 0;
lock (g1) { c1 = 3; }
lock (g2) { c2 = 4; }
__P((c1 + c2).ToString());
__Check("7");

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
