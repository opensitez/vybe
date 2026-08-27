// vybe-test: csharp/csharp_namespace_aliases/qualified_reference_accesses_nested_enum_member
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

__P((Demo.Job.State.Done).ToString());
__Check("Done");

namespace Demo { public class Job { public enum State { Idle, Done } } }

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
