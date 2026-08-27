// vybe-test: csharp/csharp_object_equality/null_equals_null_returns_true_via_static_method
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

__P((object.Equals(null, null)).ToString());
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
