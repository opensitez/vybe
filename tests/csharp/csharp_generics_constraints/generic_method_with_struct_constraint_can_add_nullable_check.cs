// vybe-test: csharp/csharp_generics_constraints/generic_method_with_struct_constraint_can_add_nullable_check
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

__P((Describe<int>(7)).ToString());
__Check("7");

string Describe<T>(T? value) where T : struct { return value.HasValue ? value.Value.ToString() : "none"; }

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
