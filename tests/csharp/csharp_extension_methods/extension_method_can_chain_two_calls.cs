// vybe-test: csharp/csharp_extension_methods/extension_method_can_chain_two_calls
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

__P(("x".AddA().AddB()).ToString());
__Check("xab");

namespace Demo { public static class TextExt { public static string AddA(this string value) { return value + "a"; } public static string AddB(this string value) { return value + "b"; } } }

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
