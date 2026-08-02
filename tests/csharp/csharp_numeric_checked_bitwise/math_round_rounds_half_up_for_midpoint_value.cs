// vybe-test: csharp/csharp_numeric_checked_bitwise/math_round_rounds_half_up_for_midpoint_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Round(4.5)).ToString(), "4");
