// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

using static __Harness;

// list_initialization_surface
int seed = 30;
__P((seed + 1 > seed).ToString());
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
