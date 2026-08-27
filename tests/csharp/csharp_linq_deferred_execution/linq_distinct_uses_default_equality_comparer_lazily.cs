// vybe-test: csharp/csharp_linq_deferred_execution/linq_distinct_uses_default_equality_comparer_lazily
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using static __Harness;
using System.Linq;

foreach (var value in new[] { 1, 1, 2, 2, 3 }.Distinct()) __P((value).ToString());
__Check("1\n2\n3");

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
