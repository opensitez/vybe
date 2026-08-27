// vybe-test: csharp/csharp_tuples_advanced/tuple_equality_compares_element_wise
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

using static __Harness;

var a = (1, "x");
var b = (1, "x");
__P((a == b).ToString());
__Check("True");

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
