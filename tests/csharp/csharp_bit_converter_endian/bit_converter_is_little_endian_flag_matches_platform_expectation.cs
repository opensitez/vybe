// vybe-test: csharp/csharp_bit_converter_endian/bit_converter_is_little_endian_flag_matches_platform_expectation
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.BitConverter.IsLittleEndian).ToString(), "True");
