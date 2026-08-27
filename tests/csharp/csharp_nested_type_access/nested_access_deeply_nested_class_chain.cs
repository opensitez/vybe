// vybe-test: csharp/csharp_nested_type_access/nested_access_deeply_nested_class_chain
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new L1.L2.L3().V).ToString());
__Check("42");

class L1{public class L2{public class L3{public int V=42;}}}

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
