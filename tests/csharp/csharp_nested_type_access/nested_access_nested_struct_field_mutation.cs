// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_field_mutation
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Canvas().Make().X).ToString());
__Check("9");

class Canvas{public struct Dot{public int X;} public Dot Make(){var d=new Dot(); d.X=9; return d;}}

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
