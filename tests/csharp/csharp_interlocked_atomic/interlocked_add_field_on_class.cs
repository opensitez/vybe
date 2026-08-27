// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

var c = new Counter();
c.Add(3);
__P((c.Value).ToString());
__Check("8");

class Counter {
    public int Value = 5;
    public void Add(int n) { System.Threading.Interlocked.Add(ref Value, n); }
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
