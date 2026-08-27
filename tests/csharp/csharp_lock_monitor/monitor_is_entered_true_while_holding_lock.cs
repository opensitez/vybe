// vybe-test: csharp/csharp_lock_monitor/monitor_is_entered_true_while_holding_lock
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object gate = new object();
int count = 0;
System.Threading.Monitor.Enter(gate);
count = System.Threading.Monitor.IsEntered(gate) ? 1 : 0;
System.Threading.Monitor.Exit(gate);
__P((count).ToString());
__Check("1");

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
