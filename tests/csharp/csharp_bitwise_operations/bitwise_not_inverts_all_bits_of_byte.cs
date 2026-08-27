// vybe-test: csharp/csharp_bitwise_operations/bitwise_not_inverts_all_bits_of_byte
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

using static __Harness;

byte b = 0b11110000;
__P(((byte)(~b)).ToString());
__Check("15");

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
