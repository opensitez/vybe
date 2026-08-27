// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

using static __Harness;

// explicit_typing_surface
int seed = 60;
__P(((seed * 2) / 2 == seed || seed == 0).ToString());
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
