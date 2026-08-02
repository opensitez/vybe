// vybe-test: csharp/csharp_numeric_checked_bitwise/compound_bitwise_or_assignment_merges_mask_bits
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 4; value |= 3; __Check((value).ToString(), "7");
