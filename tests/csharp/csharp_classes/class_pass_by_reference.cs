// vybe-test: csharp/csharp_classes/class_pass_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

void Modify(Box b) {
    b.Value = 99;
}
var b = new Box();
b.Value = 1;
Modify(b);
__P((b.Value).ToString());
__Check("99");

class Box {
    public int Value;
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
