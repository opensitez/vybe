// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_field_holds_nested_struct
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Grid().Read()).ToString());
__Check("6");

class Grid{public struct Cell{public int V;} Cell _c; public Grid(){_c.V=6;} public int Read()=>_c.V;}

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
