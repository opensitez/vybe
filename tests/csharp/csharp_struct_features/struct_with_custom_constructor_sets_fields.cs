// vybe-test: csharp/csharp_struct_features/struct_with_custom_constructor_sets_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

var r = new Rect(3, 4);
__P((r.W * r.H).ToString());
__Check("12");

struct Rect { public int W,H; public Rect(int w, int h) { W=w; H=h; } }

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
