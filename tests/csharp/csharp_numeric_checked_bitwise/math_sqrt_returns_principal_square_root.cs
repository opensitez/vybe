// vybe-test: csharp/csharp_numeric_checked_bitwise/math_sqrt_returns_principal_square_root
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Sqrt(81)).ToString(), "9");
