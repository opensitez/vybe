// vybe-test: csharp/csharp_comparison_sorting/order_by_with_key_projection_sorts_by_length
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Linq;

var values = new[] { "bbb", "a", "cc" }
.OrderBy(text => text.Length);
foreach (var value in values) __P((value).ToString());
__Check("a\ncc\nbbb");

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
