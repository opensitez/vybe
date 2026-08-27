// vybe-test: csharp/csharp_namespace_aliases/using_directive_imports_custom_namespace_for_unqualified_access
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;
using Demo.Tools;

__P((new Worker().Name).ToString());
__Check("tool");

namespace Demo.Tools { public class Worker { public string Name => "tool"; } }

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
