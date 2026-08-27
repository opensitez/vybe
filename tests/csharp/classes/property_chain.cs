// vybe-test: csharp/classes/property_chain
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var o = new Outer(42);
__P((o.inner.value).ToString());
__Check("42");

class Inner { public int value; public Inner(int v) { this.value = v; } }

class Outer { public Inner inner; public Outer(int v) { this.inner = new Inner(v); } }

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
