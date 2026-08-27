// vybe-test: csharp/csharp_number_bases/long_hex_literal_covers_full_64_bit_range
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

using static __Harness;

__P((0x7FFFFFFFFFFFFFFFL==long.MaxValue).ToString());
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
