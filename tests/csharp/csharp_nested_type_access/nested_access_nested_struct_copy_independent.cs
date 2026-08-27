// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_copy_independent
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Sheet().Sum()).ToString());
__Check("8");

class Sheet{public struct Cell{public int V;} public int Sum(){var a=new Cell(); var b=a; a.V=3; b.V=5; return a.V+b.V;}}

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
