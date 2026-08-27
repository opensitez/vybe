// vybe-test: csharp/csharp_comparison_sorting/sequence_equal_reports_false_for_different_arrays
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Linq;

__P((new[] { 1, 2 }.SequenceEqual(new[] { 2, 1 })).ToString());
__Check("False");

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
