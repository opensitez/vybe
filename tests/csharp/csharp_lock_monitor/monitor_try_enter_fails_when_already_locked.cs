// vybe-test: csharp/csharp_lock_monitor/monitor_try_enter_fails_when_already_locked
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object lk = new object();
int entered = 0;
bool lockTaken = false;
System.Threading.Monitor.Enter(lk, ref lockTaken);
if (lockTaken) {
    entered = 1;
    System.Threading.Monitor.Exit(lk);
}
__P(entered.ToString());
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
