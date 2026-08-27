// vybe-test: csharp/csharp_covariance_contravariance/array_covariance_allows_derived_array_in_base_array_reference
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

using static __Harness;

string[] strings = { "a", "b" }
;
object[] objects = strings;
__P((objects[0]).ToString());
__Check("a");

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
