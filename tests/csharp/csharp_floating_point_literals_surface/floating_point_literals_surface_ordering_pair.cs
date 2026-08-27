// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

using static __Harness;

// floating_point_literals_surface
int seed = 16;
int right = seed + 1;
__P((seed < right).ToString());
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
