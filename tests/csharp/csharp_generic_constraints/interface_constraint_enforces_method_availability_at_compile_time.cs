// vybe-test: csharp/csharp_generic_constraints/interface_constraint_enforces_method_availability_at_compile_time
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

string Get<T>(T t) where T : ILabel => t.Label();
__P((Get(new Tag())).ToString());
__Check("tag");

interface ILabel { string Label(); }

class Tag : ILabel { public string Label() => "tag"; }

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
