// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_can_access_properties_of_extended_type
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

using static __Harness;

__P((new Box{Width=3,Height=4}.Area()).ToString());
__Check("12");

class Box { public int Width, Height; }

static class BoxExt { public static int Area(this Box b) => b.Width*b.Height; }

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
