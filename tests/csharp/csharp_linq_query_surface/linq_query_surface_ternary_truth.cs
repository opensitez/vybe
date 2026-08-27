// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

using static __Harness;

// linq_query_surface
int seed = 117;
bool cond = seed % 2 == 0;
__P((cond || !cond).ToString());
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
