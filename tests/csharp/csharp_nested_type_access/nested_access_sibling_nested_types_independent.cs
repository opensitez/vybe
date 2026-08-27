// vybe-test: csharp/csharp_nested_type_access/nested_access_sibling_nested_types_independent
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Duo.A().Bump(5)).ToString());
__P((new Duo.B().Bump(5)).ToString());
__Check("6\n7");

class Duo{public class A{public int Bump(int n)=>n+1;} public class B{public int Bump(int n)=>n+2;}}

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
