// vybe-test: csharp/csharp_generics_constraints/generic_method_with_class_constraint_accepts_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

__P((Echo("text")).ToString());
__Check("text");

string Echo<T>(T value) where T : class { return value.ToString(); }

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
