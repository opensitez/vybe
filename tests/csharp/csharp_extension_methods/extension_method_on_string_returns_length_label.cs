// vybe-test: csharp/csharp_extension_methods/extension_method_on_string_returns_length_label
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

__P(("abc".Label()).ToString());
__Check("abc:3");

namespace Demo { public static class TextExt { public static string Label(this string value) { return value + ":" + value.Length; } } }

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
