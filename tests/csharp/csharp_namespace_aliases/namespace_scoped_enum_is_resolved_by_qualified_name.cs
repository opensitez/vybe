// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_enum_is_resolved_by_qualified_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

__P((Demo.State.Ready).ToString());
__Check("Ready");

namespace Demo { public enum State { Ready } }

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
