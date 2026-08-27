// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

using static __Harness;

// using_disposal_surface
double seed = 52;
__P(((seed + 0.5 - 0.5) == seed).ToString());
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
