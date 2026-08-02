// vybe-test: csharp/csharp_numeric_checked_bitwise/math_floor_truncates_toward_negative_infinity
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Floor(3.9)).ToString(), "3");
