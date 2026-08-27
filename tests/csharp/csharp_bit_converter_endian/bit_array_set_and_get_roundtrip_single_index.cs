// vybe-test: csharp/csharp_bit_converter_endian/bit_array_set_and_get_roundtrip_single_index
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

using static __Harness;

var bits = new System.Collections.BitArray(3);
bits[1] = true;
__P((bits[1]).ToString());
__P((bits[0]).ToString());
__Check("True\nFalse");

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
