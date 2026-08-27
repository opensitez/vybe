// vybe-test: csharp/csharp_bit_converter_endian/bit_converter_is_little_endian_flag_matches_platform_expectation
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

using static __Harness;

__P((System.BitConverter.IsLittleEndian).ToString());
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
