// vybe-test: csharp/csharp_generics_where/where_struct_constraint_accepts_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

T Default<T>() where T:struct=>default;
__P((Default<int>()).ToString());
__P((Default<bool>()).ToString());
__Check("0\nFalse");

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
