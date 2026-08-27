// vybe-test: csharp/csharp_extension_methods/extension_method_can_compose_with_generator_output
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;
using System.Collections.Generic;

foreach (var value in new[] { 1, 2 }.Twice()) __P((value).ToString());
__Check("2\n4");

namespace Demo { public static class NumberExt { public static IEnumerable<int> Twice(this IEnumerable<int> values) { foreach (var value in values) yield return value * 2; } } }

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
