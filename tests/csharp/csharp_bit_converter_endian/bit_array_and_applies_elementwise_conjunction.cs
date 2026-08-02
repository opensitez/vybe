// vybe-test: csharp/csharp_bit_converter_endian/bit_array_and_applies_elementwise_conjunction
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = new System.Collections.BitArray(new bool[] { true, false, true });
var right = new System.Collections.BitArray(new bool[] { true, true, false });
left.And(right);
__Check((left[0]).ToString(), "True");
__Check((left[1]).ToString(), "False");
