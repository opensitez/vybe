// vybe-test: csharp/csharp_numeric_types/byte_wraps_to_zero_on_unchecked_overflow
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

unchecked { byte b = 255; b++; __P((b).ToString()); }
__Check("0");

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
