// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

using static __Harness;

// partial_type_behavior
string feature = "partial_type_behavior";
__P((feature.Substring(0, 1) == feature[0].ToString()).ToString());
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
