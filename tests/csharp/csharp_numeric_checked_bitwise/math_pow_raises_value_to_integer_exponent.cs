// vybe-test: csharp/csharp_numeric_checked_bitwise/math_pow_raises_value_to_integer_exponent
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Pow(2, 5)).ToString(), "32");
