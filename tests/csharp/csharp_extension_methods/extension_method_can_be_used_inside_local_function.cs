// vybe-test: csharp/csharp_extension_methods/extension_method_can_be_used_inside_local_function
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

int Read() { return 9.Inc(); }
__P((Read()).ToString());
__Check("10");

namespace Demo { public static class NumberExt { public static int Inc(this int value) { return value + 1; } } }

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
