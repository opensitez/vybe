// vybe-test: csharp/csharp_generics_advanced/default_of_generic_t_is_null_for_reference_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

__P((Null<string>() == null).ToString());
__Check("True");

T Null<T>() where T : class => default(T);

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
