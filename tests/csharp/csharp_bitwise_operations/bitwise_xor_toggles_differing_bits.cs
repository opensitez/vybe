// vybe-test: csharp/csharp_bitwise_operations/bitwise_xor_toggles_differing_bits
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0b1010 ^ 0b1100).ToString(), "6");
