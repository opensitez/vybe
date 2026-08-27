// vybe-test: csharp/csharp_lock_monitor/lock_two_objects_independent_counters
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object a = new object();
object b = new object();
int ca = 0;
int cb = 0;
lock (a) { ca++; }
lock (b) { cb += 2; }
__P((ca + cb).ToString());
__Check("3");

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
