// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_preserves_declaration_order
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

__P((System.Enum.GetNames(typeof(Order))[1]).ToString());
__Check("A");

enum Order{Z,A,M}

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
