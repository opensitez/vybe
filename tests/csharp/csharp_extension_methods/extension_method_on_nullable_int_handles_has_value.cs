// vybe-test: csharp/csharp_extension_methods/extension_method_on_nullable_int_handles_has_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

int? value = 8;
__P((value.Describe()).ToString());
__Check("8");

namespace Demo { public static class NullableExt { public static string Describe(this int? value) { return value.HasValue ? value.Value.ToString() : "none"; } } }

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
