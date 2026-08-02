// vybe-test: csharp/csharp_bitwise_operations/signed_right_shift_preserves_sign_bit_for_negative
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((-8 >> 1).ToString(), "-4");
