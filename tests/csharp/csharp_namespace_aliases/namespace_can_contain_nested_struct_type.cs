// vybe-test: csharp/csharp_namespace_aliases/namespace_can_contain_nested_struct_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

var point = new Demo.Point { X = 2, Y = 5 }
;
__P((point.X + point.Y).ToString());
__Check("7");

namespace Demo { public struct Point { public int X; public int Y; } }

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
