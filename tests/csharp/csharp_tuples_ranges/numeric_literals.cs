// vybe-test: csharp/csharp_tuples_ranges/numeric_literals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

using static __Harness;

__P((0xFF).ToString());
__P((0b1010).ToString());
__P((1.5e2).ToString());
__Check("255\n10\n150");

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
