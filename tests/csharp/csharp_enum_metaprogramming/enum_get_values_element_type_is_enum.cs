// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_element_type_is_enum
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var values=System.Enum.GetValues(typeof(Kind));
__P((values.GetType().GetElementType().Name).ToString());
__Check("Kind");

enum Kind{A,B}

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
