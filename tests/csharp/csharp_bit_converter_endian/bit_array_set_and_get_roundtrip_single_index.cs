// vybe-test: csharp/csharp_bit_converter_endian/bit_array_set_and_get_roundtrip_single_index
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bits = new System.Collections.BitArray(3);
bits[1] = true;
__Check((bits[1]).ToString(), "True");
__Check((bits[0]).ToString(), "False");
