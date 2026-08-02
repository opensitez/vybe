// vybe-test: csharp/csharp_numeric_checked_bitwise/unchecked_block_wraps_byte_overflow_without_throwing
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unchecked { byte value = 255; value += 1; __Check((value).ToString(), "0"); }
