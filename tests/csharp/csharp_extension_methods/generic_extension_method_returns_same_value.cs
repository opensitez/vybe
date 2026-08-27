// vybe-test: csharp/csharp_extension_methods/generic_extension_method_returns_same_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

__P(("hi".Echo()).ToString());
__Check("hi");

namespace Demo { public static class EchoExt { public static T Echo<T>(this T value) { return value; } } }

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
