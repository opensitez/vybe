// vybe-test: csharp/type_gettype_dynamic/get_type_resolves_a_user_declared_class
// origin: languages/csharp/tests/csharp/test_type_gettype_dynamic.rs

using static __Harness;

__P("Valid_get_type_resolves_a_user_declared_class");
__Check("Valid_get_type_resolves_a_user_declared_class");
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
