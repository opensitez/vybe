// vybe-test: csharp/csharp_generic_constraints/class_constraint_allows_null_assignment_to_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

__P((AsNull<string>() == null).ToString());
__Check("True");

T AsNull<T>() where T : class => null;

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
