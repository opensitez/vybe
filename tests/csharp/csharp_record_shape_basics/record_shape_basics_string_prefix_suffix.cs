// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

using static __Harness;

// record_shape_basics
string feature = "record_shape_basics";
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
