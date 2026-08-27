// vybe-test: csharp/csharp_generics_advanced/typeof_on_closed_generic_includes_type_arg_in_name
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

__P((typeof(System.Collections.Generic.List<int>).IsGenericType).ToString());
__Check("True");

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
