// vybe-test: csharp/csharp_math_advanced/math_round_midpoint_away_from_zero
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Round(2.5,System.MidpointRounding.AwayFromZero)).ToString(), "3");
