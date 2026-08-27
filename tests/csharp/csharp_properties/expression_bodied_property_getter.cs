// vybe-test: csharp/csharp_properties/expression_bodied_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

__P((new Rect{W=3,H=4}.Area).ToString());
__Check("12");

class Rect { public int W,H; public int Area => W * H; }

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
