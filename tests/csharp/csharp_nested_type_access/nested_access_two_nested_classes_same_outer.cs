// vybe-test: csharp/csharp_nested_type_access/nested_access_two_nested_classes_same_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Pair.Left().V).ToString());
__P((new Pair.Right().V).ToString());
__Check("1\n2");

class Pair{public class Left{public int V=1;} public class Right{public int V=2;}}

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
