// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_in_partial_outer_part
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Worker().Go()).ToString());
__Check("1");

partial class Worker{public class Helper{public int Run()=>1;}}

partial class Worker{public int Go()=>new Helper().Run();}

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
