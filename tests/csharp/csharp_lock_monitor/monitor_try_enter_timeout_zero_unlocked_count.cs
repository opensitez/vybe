// vybe-test: csharp/csharp_lock_monitor/monitor_try_enter_timeout_zero_unlocked_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

object gate = new object();
bool got = System.Threading.Monitor.TryEnter(gate);
int count = got ? 1 : 0;
if (got) System.Threading.Monitor.Exit(gate);
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
