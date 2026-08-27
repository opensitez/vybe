// vybe-test: csharp/linq_lambdas/event_with_args
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var t = new Timer();
t.OnTick += n => __P(("tick " + n).ToString());
t.Tick(1);
t.Tick(2);
__Check("tick 1\ntick 2");

class Timer {
    public event Action<int> OnTick;
    public void Tick(int count) { if (OnTick != null) OnTick(count); }
}

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
