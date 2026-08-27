// vybe-test: csharp/csharp_lock_monitor/lock_task_run_two_gates_isolated_totals
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object g1 = new object();
object g2 = new object();
int c1 = 0;
int c2 = 0;
var t1 = System.Threading.Tasks.Task.Run(() => { lock (g1) { c1 += 5; } });
var t2 = System.Threading.Tasks.Task.Run(() => { lock (g2) { c2 += 6; } });
System.Threading.Tasks.Task.WaitAll(t1, t2);
__P((c1 + c2).ToString());
__Check("11");

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
