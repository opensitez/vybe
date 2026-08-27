// vybe-test: csharp/csharp_nested_classes/deeply_nested_class_visible_through_chain
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

__P((new A.B.C().V).ToString());
__Check("3");

class A{public class B{public class C{public int V=3;}}}

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
