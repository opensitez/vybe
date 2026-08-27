// vybe-test: csharp/csharp_extension_methods/extension_method_with_generic_receiver_can_count_items
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;
using System.Collections.Generic;

__P((new[] { 1, 2, 3 }.CountItems()).ToString());
__Check("3");

namespace Demo { public static class EnumerableExt { public static int CountItems<T>(this IEnumerable<T> items) { int total = 0; foreach (var _ in items) total++; return total; } } }

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
