// vybe-test: csharp/csharp_extension_methods/extension_method_can_use_namespace_import_from_nested_namespace
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo.Tools;

__P(("go".Bang()).ToString());
__Check("go!");

namespace Demo.Tools { public static class TextExt { public static string Bang(this string value) { return value + "!"; } } }

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
