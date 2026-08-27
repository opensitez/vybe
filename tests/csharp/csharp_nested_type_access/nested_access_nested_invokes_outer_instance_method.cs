// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_invokes_outer_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Outer().Via(6)).ToString());
__Check("12");

class Outer{int Double(int n)=>n*2; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Run(int n)=>o.Double(n);} public int Via(int n)=>new Inner(this).Run(n);}

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
