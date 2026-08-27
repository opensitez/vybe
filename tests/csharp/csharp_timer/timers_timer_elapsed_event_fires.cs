// vybe-test: csharp/csharp_timer/timers_timer_elapsed_event_fires
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

using static __Harness;

bool fired=false;
var t=new System.Timers.Timer(10){AutoReset=false}
;
t.Elapsed+=(_,__)=>fired=true;
t.Start();
System.Threading.Thread.Sleep(100);
__P((fired).ToString());
__Check("True");

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
