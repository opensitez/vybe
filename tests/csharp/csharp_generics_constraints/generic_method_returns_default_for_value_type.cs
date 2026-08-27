// vybe-test: csharp/csharp_generics_constraints/generic_method_returns_default_for_value_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

__P((Zero<int>()).ToString());
__Check("0");

T Zero<T>() where T : struct { return default(T); }

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
