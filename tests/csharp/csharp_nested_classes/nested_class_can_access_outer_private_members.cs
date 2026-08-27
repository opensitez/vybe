// vybe-test: csharp/csharp_nested_classes/nested_class_can_access_outer_private_members
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

__P((new Outer.Inner().Get()).ToString());
__Check("42");

class Outer{
    static int secret=42;
    public class Inner{public int Get()=>secret;}
}

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
