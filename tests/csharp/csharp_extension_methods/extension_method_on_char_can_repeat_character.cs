// vybe-test: csharp/csharp_extension_methods/extension_method_on_char_can_repeat_character
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;

__P(('x'.Repeat(3)).ToString());
__Check("xxx");

namespace Demo { public static class CharExt { public static string Repeat(this char value, int count) { return new string(value, count); } } }

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
