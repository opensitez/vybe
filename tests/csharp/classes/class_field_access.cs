// vybe-test: csharp/classes/class_field_access
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var b = new Box(42);
__P((b.value).ToString());
__Check("42");

class Box {
            public int value;
            public Box(int v) { this.value = v; }
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
