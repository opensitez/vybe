// vybe-test: csharp/csharp_object_equality/reference_equals_returns_false_for_two_distinct_object_instances
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

var a = new object();
var b = new object();
__P((object.ReferenceEquals(a, b)).ToString());
__Check("False");

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
