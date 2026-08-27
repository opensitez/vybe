// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_reads_outer_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Outer().Via()).ToString());
__Check("4");

class Outer{int seed=4; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.seed;} public int Via()=>new Inner(this).Read();}

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
