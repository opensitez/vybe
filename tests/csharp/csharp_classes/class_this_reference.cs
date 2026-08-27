// vybe-test: csharp/csharp_classes/class_this_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var c = new Counter();
c.Increment();
c.Increment();
c.Increment();
__P((c.GetCount()).ToString());
__Check("3");

class Counter {
    private int count = 0;
    public void Increment() { this.count++; }
    public int GetCount() { return this.count; }
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
