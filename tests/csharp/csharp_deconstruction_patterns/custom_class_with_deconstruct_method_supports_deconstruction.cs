// vybe-test: csharp/csharp_deconstruction_patterns/custom_class_with_deconstruct_method_supports_deconstruction
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

using static __Harness;

var (w, h) = new Size{W=3,H=4}
;
__P((w).ToString());
__P((h).ToString());
__Check("3\n4");

class Size {
    public int W, H;
    public void Deconstruct(out int w, out int h) { w=W; h=H; }
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
