// vybe-test: csharp/csharp_extension_methods/extension_method_on_object_accepts_boxed_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

object value = 3;
__P((value.Kind()).ToString());
__Check("Int32");

namespace Demo { public static class ObjectExt { public static string Kind(this object value) { return value.GetType().Name; } } }

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
