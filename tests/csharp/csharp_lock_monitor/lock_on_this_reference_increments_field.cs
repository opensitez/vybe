// vybe-test: csharp/csharp_lock_monitor/lock_on_this_reference_increments_field
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

using static __Harness;

var box = new Box();
box.Inc();
__P((box.counter).ToString());
__Check("1");

class Box {
    public int counter = 0;
    public void Inc() { lock (this) { counter++; } }
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
