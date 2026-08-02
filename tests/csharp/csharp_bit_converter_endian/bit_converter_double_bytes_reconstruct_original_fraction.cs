// vybe-test: csharp/csharp_bit_converter_endian/bit_converter_double_bytes_reconstruct_original_fraction
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.BitConverter.GetBytes(2.5);
__Check((System.BitConverter.ToDouble(bytes, 0)).ToString(), "2.5");
