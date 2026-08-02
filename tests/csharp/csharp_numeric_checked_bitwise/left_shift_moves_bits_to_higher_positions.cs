// vybe-test: csharp/csharp_numeric_checked_bitwise/left_shift_moves_bits_to_higher_positions
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((3 << 2).ToString(), "12");
