// vybe-test: csharp/csharp_generic_constraints/new_constraint_allows_default_constructor_call_inside_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

T Create<T>() where T : new() => new T();
var w = Create<Widget>();
__P((w.Value).ToString());
__Check("42");

class Widget { public int Value = 42; }

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
