// vybe-test: csharp/csharp_bit_converter_endian/bit_converter_int32_roundtrip_preserves_numeric_value
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.BitConverter.GetBytes(1024);
__Check((System.BitConverter.ToInt32(bytes, 0)).ToString(), "1024");
