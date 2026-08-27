// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_invokes_outer_static_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Outer.Inner().Run(2)).ToString());
__Check("6");

class Outer{static int Triple(int n)=>n*3; public class Inner{public int Run(int n)=>Triple(n);}}

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
