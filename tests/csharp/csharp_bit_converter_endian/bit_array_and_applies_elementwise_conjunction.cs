// vybe-test: csharp/csharp_bit_converter_endian/bit_array_and_applies_elementwise_conjunction
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

using static __Harness;

var left = new System.Collections.BitArray(new bool[] { true, false, true });
var right = new System.Collections.BitArray(new bool[] { true, true, false });
left.And(right);
__P((left[0]).ToString());
__P((left[1]).ToString());
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
