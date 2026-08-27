// vybe-test: csharp/csharp_classes/class_multiple_instances
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var a = new Box(10);
var b = new Box(20);
__P((a.Value + b.Value).ToString());
__Check("30");

class Box {
    public int Value;
    public Box(int v) { Value = v; }
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
