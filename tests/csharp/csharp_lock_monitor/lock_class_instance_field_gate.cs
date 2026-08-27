// vybe-test: csharp/csharp_lock_monitor/lock_class_instance_field_gate
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

var sc = new SafeCounter();
sc.Add(3);
sc.Add(4);
__P((sc.Value).ToString());
__Check("7");

class SafeCounter {
    private object gate = new object();
    public int Value = 0;
    public void Add(int n) { lock (gate) { Value += n; } }
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
