// vybe-test: csharp/csharp_numeric_formatting/format_g_chooses_shorter_of_fixed_or_scientific_representation
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

using static __Harness;

__P(((0.00001).ToString("G")).ToString());
__Check("1E-05");

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
