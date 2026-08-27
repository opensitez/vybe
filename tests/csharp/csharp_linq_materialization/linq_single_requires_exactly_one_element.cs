// vybe-test: csharp/csharp_linq_materialization/linq_single_requires_exactly_one_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

using static __Harness;
using System.Linq;

__P((new[] { 42 }.Single()).ToString());
__Check("42");

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
