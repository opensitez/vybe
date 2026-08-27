// vybe-test: csharp/csharp_extension_methods/extension_method_on_array_can_sum_elements
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

__P((new[] { 1, 2, 3 }.SumAll()).ToString());
__Check("6");

namespace Demo { public static class ArrayExt { public static int SumAll(this int[] values) { int total = 0; foreach (var value in values) total += value; return total; } } }

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
