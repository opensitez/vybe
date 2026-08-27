// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_captures_outer_const
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Outer.Inner().Tag()).ToString());
__Check("prefix");

class Outer{public const string Prefix="pre"; public class Inner{public string Tag()=>Prefix+"fix";}}

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
