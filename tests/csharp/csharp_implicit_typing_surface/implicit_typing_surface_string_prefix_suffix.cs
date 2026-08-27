// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

using static __Harness;

// implicit_typing_surface
string feature = "implicit_typing_surface";
__P((feature.Substring(0, 1) == feature[0].ToString()).ToString());
__Check("True");

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
