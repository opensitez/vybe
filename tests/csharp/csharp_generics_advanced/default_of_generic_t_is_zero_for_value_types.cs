// vybe-test: csharp/csharp_generics_advanced/default_of_generic_t_is_zero_for_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

T Zero<T>() => default(T);
__P((Zero<int>()).ToString());
__P((Zero<bool>()).ToString());
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
