// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_shorten_fully_qualified_type_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;
using Thing = Demo.Tools.Box;

__P((new Thing().Value).ToString());
__Check("7");

namespace Demo.Tools { public class Box { public int Value = 7; } }

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
