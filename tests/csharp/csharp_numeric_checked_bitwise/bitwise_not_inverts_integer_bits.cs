// vybe-test: csharp/csharp_numeric_checked_bitwise/bitwise_not_inverts_integer_bits
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((~5).ToString(), "-6");
