// vybe-test: csharp/csharp_namespace_aliases/global_system_namespace_type_is_available_inside_custom_namespace
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

__P((new Demo.Worker().Read()).ToString());
__Check("a,b");

namespace Demo { public class Worker { public string Read() { return global::System.String.Join(",", new[] { "a", "b" }); } } }

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
