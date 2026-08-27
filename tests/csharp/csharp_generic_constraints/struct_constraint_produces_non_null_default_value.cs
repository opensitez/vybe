// vybe-test: csharp/csharp_generic_constraints/struct_constraint_produces_non_null_default_value
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

__P((Zero<int>()).ToString());
__Check("0");

T Zero<T>() where T : struct => default;

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
