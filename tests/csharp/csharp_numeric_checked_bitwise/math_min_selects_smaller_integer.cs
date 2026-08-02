// vybe-test: csharp/csharp_numeric_checked_bitwise/math_min_selects_smaller_integer
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Min(4, 7)).ToString(), "4");
