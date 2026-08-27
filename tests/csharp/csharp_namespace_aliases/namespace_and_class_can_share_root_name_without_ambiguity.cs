// vybe-test: csharp/csharp_namespace_aliases/namespace_and_class_can_share_root_name_without_ambiguity
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

__P((new Demo.Sub.Demo().Value).ToString());
__Check("5");

namespace Demo.Sub { public class Demo { public int Value = 5; } }

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
