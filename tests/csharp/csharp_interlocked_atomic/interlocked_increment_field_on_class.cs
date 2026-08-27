// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

var c = new Counter();
c.Bump();
c.Bump();
__P((c.Value).ToString());
__Check("2");

class Counter {
    public int Value = 0;
    public void Bump() { System.Threading.Interlocked.Increment(ref Value); }
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
