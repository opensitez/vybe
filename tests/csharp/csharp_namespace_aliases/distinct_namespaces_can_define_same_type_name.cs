// vybe-test: csharp/csharp_namespace_aliases/distinct_namespaces_can_define_same_type_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

__P((new Left.Item().Name).ToString());
__P((new Right.Item().Name).ToString());
__Check("L\nR");

namespace Left { public class Item { public string Name => "L"; } }

namespace Right { public class Item { public string Name => "R"; } }

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
