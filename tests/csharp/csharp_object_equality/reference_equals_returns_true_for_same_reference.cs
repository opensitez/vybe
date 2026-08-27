// vybe-test: csharp/csharp_object_equality/reference_equals_returns_true_for_same_reference
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

var a = new object();
var b = a;
__P((object.ReferenceEquals(a, b)).ToString());
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
