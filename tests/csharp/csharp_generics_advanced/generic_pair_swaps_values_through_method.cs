// vybe-test: csharp/csharp_generics_advanced/generic_pair_swaps_values_through_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

(T, T) Swap<T>(T a, T b) => (b, a);
var (x, y) = Swap(1, 2);
__P((x).ToString());
__P((y).ToString());
__Check("2\n1");

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
