// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_exposes_nested_via_property
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Shell().Inner.Id).ToString());
__Check("2");

class Shell{public class Core{public int Id=2;} Core _c=new Core(); public Core Inner=>_c;}

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
