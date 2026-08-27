// vybe-test: csharp/classes/class_multiple_instances
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var a = new Counter(0);
var b = new Counter(100);
a.Inc();
a.Inc();
b.Inc();
__P((a.Get()).ToString());
__P((b.Get()).ToString());
__Check("2\n101");

class Counter {
            int count;
            public Counter(int start) { this.count = start; }
            public void Inc() { this.count = this.count + 1; }
            public int Get() { return this.count; }
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
