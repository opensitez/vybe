// vybe-test: csharp/csharp_conversion_methods/convert_to_int32_from_string
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

using static __Harness;

// conversion_methods
__P((System.Convert.ToInt32("99")).ToString());
__Check("99");

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
