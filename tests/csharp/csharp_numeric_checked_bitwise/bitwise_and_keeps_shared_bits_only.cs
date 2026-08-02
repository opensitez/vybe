// vybe-test: csharp/csharp_numeric_checked_bitwise/bitwise_and_keeps_shared_bits_only
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((6 & 3).ToString(), "2");
